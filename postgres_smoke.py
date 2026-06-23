"""Destructive database smoke test; refuses to run outside an explicit test profile."""
from __future__ import annotations

import os
import time

import pandas as pd


def main():
    database_url = os.environ.get("DATABASE_URL", "")
    data_dir = os.environ.get("BOQ_DATA_DIR", "")
    target_label = database_url or data_dir
    if not target_label or not any(token in target_label.lower() for token in ("test", "smoke")):
        raise SystemExit("DATABASE_URL or BOQ_DATA_DIR must identify an explicit test profile.")

    import audit
    import auth
    import db
    import repo
    import reports

    init_conn = db.init_db()
    init_conn.close()
    db.set_audit_actor({
        "UserID": None, "Username": "postgres-smoke", "DisplayName": "PostgreSQL Smoke Test"
    })
    created_projects, created_items, created_users = [], [], []
    stamp = str(int(time.time()))
    try:
        projects = repo.list_projects()
        if projects.empty:
            raise RuntimeError("Smoke test needs at least one migrated offer.")
        source_pid = int(projects.iloc[0]["ProjectID"])
        systems = repo.list_systems(source_pid)
        grid = repo.load_project_grid(source_pid, systems[0] if systems else None)
        if grid.empty:
            raise RuntimeError("Smoke test source offer has no lines.")
        suffix = repo.base_name(systems[0] if systems else "BOQ LCS").replace("BOQ", "").strip()

        pid = repo.save_offer(
            name=f"PostgreSQL Smoke {stamp}", client="Test", contact="Test",
            offer_no=f"PG-TEST-{stamp}", system_suffix=suffix or "LCS", grid=grid,
            option_label="Smoke",
        )
        created_projects.append(pid)
        loaded = repo.load_project_grid(pid)
        if not pd.Series(loaded["Margin x"]).equals(pd.Series(grid["Margin x"])):
            raise AssertionError("Saved margins did not round-trip exactly.")

        repo.update_offer(
            pid, grid, system_suffix=suffix or "LCS", option_label="Smoke Updated",
            header={
                "project": f"PostgreSQL Smoke Updated {stamp}", "client": "Test",
                "contact": "Test", "phone": "", "contractor": "", "region": "",
                "sales": "", "presales": "", "pm": "",
            },
        )
        if repo.project_meta(pid).get("OptionLabel") != "Smoke Updated":
            raise AssertionError("Offer update did not persist.")

        rev_pid, _, _ = repo.save_revision(pid, grid, system_suffix=suffix or "LCS")
        created_projects.append(rev_pid)
        opt_pid, _, _ = repo.save_option(pid, grid, "Smoke Alternative",
                                         system_suffix=suffix or "LCS")
        created_projects.append(opt_pid)

        repo.save_finance(
            pid,
            [{"Description": "Smoke invoice", "Amount (SAR)": 100, "Invoice #": "TEST-INV"}],
            [{"Description": "Smoke purchase", "Cost (SAR)": 40, "PO #": "TEST-PO"}],
        )
        payments, purchases = repo.get_finance(pid)
        if len(payments) != 1 or len(purchases) != 1:
            raise AssertionError("Finance rows did not persist.")

        tracking = repo.load_tracking(pid)
        if not tracking.empty:
            line = tracking.iloc[0]
            repo.update_tracking([(
                int(line["LineID"]), True, True, False, "TEST-PO", "", "", "", "",
                float(line["Qty"] or 0), 0,
            )])
            updated_tracking = repo.load_tracking(pid)
            if not bool(updated_tracking.iloc[0]["Paid"]):
                raise AssertionError("Tracking update did not persist.")

        repo.approve_offer(pid)
        if not repo.project_meta(pid).get("Approved"):
            raise AssertionError("Approval did not persist.")
        repo.unapprove_offer(pid)
        repo.archive_project(pid)
        if not repo.project_meta(pid).get("Archived"):
            raise AssertionError("Archive did not persist.")
        repo.unarchive_project(pid)

        item_id = repo.add_catalog_item(
            f"Smoke Brand {stamp}", f"Smoke Model {stamp}", "Smoke Description", 10, 5
        )
        created_items.append(item_id)
        if not repo.update_catalog_item(item_id, {"Description": "Smoke Updated Description"}):
            raise AssertionError("Catalogue update failed.")

        user_id = auth.create_user(f"pg_smoke_{stamp}", "temporary-test-password", "PG Smoke")
        created_users.append(user_id)
        if not any(user["UserID"] == user_id for user in auth.list_users()):
            raise AssertionError("User insert failed.")

        events, total = audit.query_events(search="postgres-smoke", limit=50)
        if total < 1:
            raise AssertionError("PostgreSQL audit trigger did not record writes.")
        if reports.offers_df().empty:
            raise AssertionError("Reports returned no migrated offers.")
        print(
            f"{('PostgreSQL' if db.is_postgres() else 'SQLite')} smoke passed: "
            f"projects={len(projects):,}, "
            f"grid_rows={len(grid):,}, audit_events={total:,}"
        )
    finally:
        if created_projects:
            repo.delete_projects(created_projects)
        if created_items:
            repo.delete_catalog_items(created_items)
        for user_id in created_users:
            auth.delete_user(user_id)


if __name__ == "__main__":
    main()
