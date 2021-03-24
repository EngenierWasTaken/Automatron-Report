
from datetime import datetime, timedelta
import locale
class BodyCsv:
    """
    Classe che data una lista di dati csv di zoom li trasforma in oggetti Python
    CSV
    """
    def __init__(self, csv):
        try:
            locale.setlocale(locale.LC_ALL, 'en_US.utf8')
            self.CsvSetter(csv)
        except Exception as e:
            print("errori nel settare BodyCsv: Errore__:", e)
            self.Nome = ""
            self.Email = ""
            self.JoinDate = ""
            self.OraEntrata = ""
            self.LeaveDate = ""
            self.OrarioUscita = ""
            self.TempoOnline = ""
            self.IngressoNormalizzato = ""
            self.UscitaNormalizzato = ""

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
    def __hash__(self):
        return hash(self.Email)
    def CsvSetter(self, csv):
        """
        funzione che setta le variabili della classe da una riga del csv
        :param csv riga del csv:
        :return:
        """
        self.Nome = csv[0]
        self.Email = csv[1]
        self.JoinDate = self.SetData(csv[2])
        self.OraEntrata = self.SetOra(csv[2])
        self.LeaveDate = self.SetData(csv[3])
        self.OrarioUscita = self.SetOra(csv[3])
        self.TempoOnline = int(csv[4])
        self.IngressoNormalizzato = self.SetOraNormalizzato(csv[2])
        self.UscitaNormalizzato = self.SetOraNormalizzato(csv[3])
        self.PrimoIngresso = None
        self.UltimoIngresso = None

    def SetData(self, Data):
        data_finale = self.GuessDateFormat(Data)
        data_finale = data_finale.strftime("%d/%m/%Y")
        data_finale = datetime.strptime(data_finale, "%d/%m/%Y")
        return data_finale

    def SetOra(self, Orario):
        data_finale = self.GuessDateFormat(Orario)
        data_finale = data_finale.strftime("%H:%M:%S")
        data_finale = datetime.strptime(data_finale, "%H:%M:%S").time()
        return data_finale

    def SetOraNormalizzato(self, Orario):
        data_finale = self.GuessDateFormat(Orario)
        data_finale = data_finale.strftime("%H:%M")
        data_finale = datetime.strptime(data_finale, "%H:%M").time()
        return data_finale

    @staticmethod
    def DividiPerEmail(BodyList):
        main_list = []
        sub_list = []
        seen = []
        for x in range(0, len(BodyList)):
            if hash(BodyList[x].Email) not in seen:
                for y in range(x, len(BodyList)):
                    if hash(BodyList[x].Email) == hash(BodyList[y].Email):
                        sub_list.append(BodyList[y])
                        seen.append(hash(BodyList[y].Email))
                main_list.append(sub_list)
                sub_list = []
        return main_list

    @staticmethod
    def bubble_sort(array):
        n = len(array)
        for i in range(n):
            already_sorted = True
            for j in range(n - i - 1):
                if array[j] > array[j + 1]:
                    array[j], array[j + 1] = array[j + 1], array[j]
                    already_sorted = False

            if already_sorted:
                break
        return array

    @staticmethod
    def SetPrimoUltimoIngresso(BodyList):

        # for x in range(0, len(BodyList)):
        #     for y in range(1, len(BodyList)):
        #         print(BodyList[y].Nome)
        #         print(BodyList[y].OraEntrata, BodyList[y].OrarioUscita)
        #         if BodyList[y].OraEntrata > BodyList[x].OraEntrata:
        #             tmp_list = BodyList[x]
        #             BodyList[x] = BodyList[y]
        #             BodyList[y] = tmp_list
        #         if BodyList[y].OrarioUscita < BodyList[x].OrarioUscita:
        #             tmp_list = BodyList[x]
        #             BodyList[x] = BodyList[y]
        #             BodyList[y] = tmp_list
        n = len(BodyList)
        for i in range(n):
            for j in range(n - i - 1):
                if BodyList[j].OraEntrata > BodyList[j + 1].OraEntrata:
                    BodyList[j], BodyList[j + 1] = BodyList[j + 1], BodyList[j]
        tmp_entrata = []
        tmp_uscita = []
        for x in BodyList:
            tmp_entrata.append(x.OraEntrata)
            tmp_uscita.append(x.OrarioUscita)

        BodyList[0].PrimoIngresso = min(tmp_entrata)
        BodyList[0].UltimoIngresso = max(tmp_uscita)


    @staticmethod
    def CalcolaTempoOnline(primo_ingresso, ultimo_ingresso):
        timedelta_primoingresso = timedelta(hours=primo_ingresso.hour, minutes=primo_ingresso.minute)
        timedelta_ultimo_ingresso = timedelta(hours=ultimo_ingresso.hour, minutes=ultimo_ingresso.minute)
        tempo_online =  timedelta_ultimo_ingresso - timedelta_primoingresso

        hour = tempo_online.seconds // 3600
        minutes = tempo_online.seconds // 60 % 60
        if minutes >= 10:
            minutes = str(minutes)
        else:
            minutes = "0"+str(minutes)


        tempo_finale = "{}:{}".format(hour, minutes)

        return tempo_finale

    @staticmethod
    def NormalizzaQuartoDora(tempo_online):
        if len(tempo_online) == 5:
            tempo = tempo_online[3:]
        else:
            tempo = tempo_online[2:]

        tempo = int(tempo)
        tempo = 15*round(tempo / 15)
        if tempo == 0:
            tempo = "00"
        if tempo == 60:
            tempo = "00"
            if len(tempo_online) == 5:
                tempo_online = int(tempo_online[:1]) + 1
                tempo_online = "0" + str(tempo_online) + ":"
            else:
                tempo_online = int(tempo_online[:1]) + 1
                tempo_online = str(tempo_online) + ":"
        if len(tempo_online) == 5:
            tempo = tempo_online[:3] + str(tempo)
        else:
            tempo = tempo_online[:2] + str(tempo)
        return tempo