import os
from Controllers.Headers import HeaderCsv
from Controllers.BodyCsv import BodyCsv
import csv
class CsvGenerator:
    """
    Legge da un file csv e crea un oggetto
        HeaderCsv
    ed Una lista di oggetti:
        BodyCsv
    e conta quante righe sono state processate
    """
    def __init__(self, csv_path):

        try:
            self.CsvReaderSetter(csv_path)
        except Exception as e:
            print("Errore durante il caricamento del file ", csv_path, "errore__ ", e)

    def CsvReaderSetter(self, csv_path):
        if os.path.exists(csv_path):
            self.BodyCsvList = []
            line_count = 0
            with open(csv_path, "r", encoding="UTF-8") as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',')
                for row in csv_reader:
                    if line_count == 0:
                        self.FieldNames_INUTILI =  (",".join(row).split(","))
                        line_count += 1
                    elif line_count == 1:
                        self.HeaderCsv = HeaderCsv(row)
                        line_count += 1
                    else:
                        try:
                            self.BodyCsvList.append(BodyCsv(row))
                            line_count += 1
                        except Exception as e:
                            print("Errore nel leggere la riga del csv", line_count, "Errore__", e)
