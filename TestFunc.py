

from Controllers.CsvReaderSetter import CsvGenerator
from Controllers.BodyCsv import BodyCsv

path = "./participants_89572547003.csv"
temp = CsvGenerator(path)
Header = temp.HeaderCsv
BodyCsvList = temp.BodyCsvList
#(int(str(x.OraEntrata[:3]).replace(":", ""))
BodyCsvList = [x for x in BodyCsvList if x.OraEntrata != ""]
SortedBodyList = sorted(BodyCsvList, key=lambda x: (x.Email, x.OraEntrata))
SortedBodyList = BodyCsv.DividiPerEmail(SortedBodyList)
#sorted_body = sorted(body_generic_list, key=lambda x: (x.email, int(str(x.primo_ingresso_normalizzato.replace(":", "")))))
print("--------- Divisore -----------")

for x in SortedBodyList:
    BodyCsv.SetPrimoUltimoIngresso(x)

for x in SortedBodyList:
    for y in x:
        print(y.Email, " Primo ", y.PrimoIngresso,"Ultimo ", y.UltimoIngresso, y.OraEntrata, y.OrarioUscita)
