
from datetime import datetime
import locale
class HeaderCsv:
    """
    Classe che dati una lista di dati csv di zoom li trasforma in oggetti Python
    Csv | Nome Tutor | Nome Docente | Orario Entrata Custom | Orario Uscita Custom
    """
    def __init__(self, csv):
        #self.NomeTutor = nomeTutor
        #self.NomeDocente = nomeDocente
        #self.OrarioEntrataCustom = orarioEntrataCustom
        #self.OrarioUscitaCustom = orarioUscitaCustom
        locale.setlocale(locale.LC_ALL, 'en_US.utf8')
        try:
            self.CsvSetter(csv)
        except Exception as e:
            print("errori nel settare HeaderCsv: Errore__:", e)

    def GuessDateFormat(self, Data):
        date_format = [
            '%d/%m/%Y %I:%M:%S %p',
            '%d/%m/%Y %H:%M:%S',
            '%m/%d/%Y %I:%M:%S %p',
            '%m/%d/%Y %H:%M:%S',
            '%Y/%m/%d %I:%M:%S %p',
            '%Y/%m/%d %H:%M:%S',
            '%Y/%d/%m %I:%M:%S %p',
            '%Y/%d/%m %H:%M:%S',
            ]
        for data in date_format:
            try:
                data_finale = datetime.strptime(Data, data)
                return data_finale
            except Exception as e:
                pass

    def CsvSetter(self, csv):
        """
        Modella la classe HeaderCsv da una row csv
        :param csv: path del csv relativo
        :return nothing
        """
        self.ID = csv[0]
        self.Argomento = csv[1]
        self.OraInizio = self.SetOra(csv[2])
        self.OraFine = self.SetOra(csv[3])
        self.Email = csv[4]
        self.Durata = csv[5]

    def SetOra(self, Orario):
        data_finale = self.GuessDateFormat(Orario)
        data_finale = data_finale.strftime("%d/%m/%Y %H:%M")
        return data_finale