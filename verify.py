import db as dbmod
c = dbmod.connect()
q = lambda s, *a: c.execute(s, a).fetchall()

print("=== PROJECTS (sample 8) ===")
for r in q("SELECT ProjectID,ProjectName,ClientName,OfferNo,CreationDate,UpdatedDate,DiscountAmount,ConversionFactor FROM Projects_Master ORDER BY ProjectID LIMIT 8"):
    print(f"  #{r['ProjectID']:>2} | {str(r['ProjectName'])[:26]:<26} | client={str(r['ClientName'])[:14]:<14} | offer={str(r['OfferNo'])[:18]:<18} | created={r['CreationDate']} updated={r['UpdatedDate']} | disc={r['DiscountAmount']} | f={r['ConversionFactor']}")

print("\n=== CONVERSION FACTORS captured (sample) ===")
for r in q("SELECT SheetName,SystemSuffix,Factor1,Factor2,Factor3 FROM Project_Sheets WHERE Factor1 IS NOT NULL LIMIT 8"):
    print(f"  {r['SheetName']:<18} suffix={str(r['SystemSuffix']):<10} f1={r['Factor1']} f2={r['Factor2']} f3={r['Factor3']}")

print("\n=== TOP CATALOGUE ITEMS by TimesQuoted ===")
for r in q("SELECT Brand,Model,substr(Description,1,34) d,UnitCostUSD,DefaultUPriceUSD,DefaultUPriceSAR,TimesQuoted FROM Items_Catalog ORDER BY TimesQuoted DESC LIMIT 12"):
    print(f"  {str(r['Brand'])[:8]:<8} {str(r['Model'])[:16]:<16} {str(r['d']):<34} cost={r['UnitCostUSD']} U$={r['DefaultUPriceUSD']} USAR={r['DefaultUPriceSAR']} x{r['TimesQuoted']}")

print("\n=== ATA Office line check (formula integrity) ===")
rows = q("""SELECT l.Description,l.Qty,l.FinalUnitCostUSD,l.TotalCostUSD,l.FinalUPriceSAR,l.TPriceSAR,l.LineType
            FROM Project_BoQ_Lines l JOIN Projects_Master p ON l.ProjectID=p.ProjectID
            WHERE p.SourceFile LIKE '%ATA Offices%' ORDER BY l.RowOrder LIMIT 6""")
for r in rows:
    qn, uc, tc = r['Qty'], r['FinalUnitCostUSD'], r['TotalCostUSD']
    chk = "OK" if (qn and uc and abs(qn*uc-(tc or 0)) < 0.01) else ("disc/svc" if r['LineType'] != 'item' else "??")
    sar = "OK" if (qn and r['FinalUPriceSAR'] and abs(qn*r['FinalUPriceSAR']-(r['TPriceSAR'] or 0)) < 0.01) else "-"
    print(f"  [{r['LineType']:<8}] {str(r['Description'])[:30]:<30} qty={qn} cost*qty={chk} SAR*qty={sar}")

print("\n=== LineType distribution ===")
for r in q("SELECT LineType,COUNT(*) n FROM Project_BoQ_Lines GROUP BY LineType"):
    print(f"  {str(r['LineType']):<10} {r['n']}")

print("\n=== Brands in catalogue ===")
for r in q("SELECT Brand,COUNT(*) n FROM Items_Catalog WHERE Brand IS NOT NULL GROUP BY Brand ORDER BY n DESC LIMIT 12"):
    print(f"  {str(r['Brand'])[:20]:<20} {r['n']}")
c.close()
