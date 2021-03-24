for x in tmp:
    x.tempo_online = self.ConvertTime(x.tempo_online)
if self.DirPicker.GetPath():
    path = self.DirPicker.GetPath()
    workbook = xlsxwriter.Workbook(f'{path}/{header_generic_list[0].ID}.xlsx')
else:
    workbook = xlsxwriter.Workbook(f'{header_generic_list[0].ID}.xlsx')


worksheet = workbook.add_worksheet(name="AttandeeReport")

cell_format_ID_riunione = workbook.add_format({"border": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
cell_format_ID_riunione_bold = workbook.add_format({"border": True, "bold": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
cell_format_catalogo = workbook.add_format({"font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
cell_format_catalogo_nome = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left"})


worksheet.set_row(0, 15)
worksheet.set_row(1, 70)
worksheet.set_row(2, 15)
worksheet.set_row(3, 35)
worksheet.set_row(4, 15)
worksheet.set_row(5, 50)
worksheet.set_row(6, 20)


worksheet.set_column(0, 0, 33.38)
worksheet.set_column(1, 1, 26.04)
worksheet.set_column(2, 2, 13.38)
worksheet.set_column(3, 3, 9.61)
worksheet.set_column(4, 4, 11.90)
worksheet.set_column(5, 5, 9.61)
worksheet.set_column(6, 6, 15.19)
worksheet.set_column(7, 7, 15.19)
worksheet.set_column(8, 8, 15.19)
worksheet.set_column(9, 9, 15.19)


worksheet.write(0, 0, "ID riunione", cell_format_ID_riunione)
worksheet.write(0, 1, "Argomento", cell_format_ID_riunione)
worksheet.write(0, 2, "Ora di inizio", cell_format_ID_riunione)
worksheet.write(0, 3, "", cell_format_ID_riunione)
worksheet.write(0, 4, "Ora di fine", cell_format_ID_riunione)
worksheet.write(0, 5, "", cell_format_ID_riunione)
worksheet.write(0, 6, "Email dell’utente", cell_format_ID_riunione)


worksheet.write(1, 0, header_generic_list[0].ID, cell_format_ID_riunione)
worksheet.write(1, 1, header_generic_list[0].Argomento, cell_format_ID_riunione_bold)
worksheet.write(1, 2, header_generic_list[0].Ora_inizio[0:10], cell_format_ID_riunione)
worksheet.write(1, 3, header_generic_list[0].Ora_inizio[10:], cell_format_ID_riunione)
worksheet.write(1, 4, header_generic_list[0].Ora_Fine[0:10], cell_format_ID_riunione)
worksheet.write(1, 5, header_generic_list[0].Ora_Fine[10:], cell_format_ID_riunione)
worksheet.write(1, 6, header_generic_list[0].Email, cell_format_ID_riunione)


tmp_primo_ingresso = datetime.strptime( self.OrarioEntrata.Value, "%H:%M").time()
tmp_ultimo_ingresso = datetime.strptime(self.OrarioUscita.Value, "%H:%M").time()
numero_ore = self.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
minutes = int(numero_ore[2:])
minutes = minutes/60
hour = int(numero_ore[:1])
orario = (hour * 60) + (minutes * 60)
worksheet.write(2, 0, "Tutor", cell_format_ID_riunione)
worksheet.write(2, 1, self.NomeTutor.Value, cell_format_ID_riunione)
worksheet.write(2, 2, "Ora inizio", cell_format_ID_riunione)
worksheet.write(2, 3, self.OrarioEntrata.Value, cell_format_ID_riunione)
worksheet.write(2, 4, "N. ore", cell_format_ID_riunione)
worksheet.write(2, 5, numero_ore, cell_format_ID_riunione)
worksheet.write(2, 6, "", cell_format_ID_riunione)


worksheet.write(3, 0, "Docente", cell_format_ID_riunione)
worksheet.write(3, 1, self.NomeDocente.Value, cell_format_ID_riunione)
worksheet.write(3, 2, "Ora Fine", cell_format_ID_riunione)
worksheet.write(3, 3, self.OrarioUscita.Value, cell_format_ID_riunione)
worksheet.write(3, 4, "Durata in min reale", cell_format_ID_riunione)
worksheet.write(3, 5, orario, cell_format_ID_riunione)
worksheet.write(3, 6, "", cell_format_ID_riunione)

worksheet.write(4, 0, "", cell_format_catalogo)
worksheet.write(4, 1, "", cell_format_catalogo)
worksheet.write(4, 2, "", cell_format_catalogo)
worksheet.write(4, 3, "", cell_format_catalogo)
worksheet.write(4, 4, "", cell_format_catalogo)
worksheet.write(4, 5, "", cell_format_catalogo)
worksheet.write(4, 6, "", cell_format_catalogo)

worksheet.write(5, 0, "Nome (nome originale)", cell_format_catalogo)
worksheet.write(5, 1, "Email dell'utente", cell_format_catalogo)
worksheet.write(5, 2, "Join Date", cell_format_catalogo)
worksheet.write(5, 3, "Ora di Entrata", cell_format_catalogo)
worksheet.write(5, 4, "Leave Date", cell_format_catalogo)
worksheet.write(5, 5, "Ora di Uscita", cell_format_catalogo)
worksheet.write(5, 6, "Durata (Minuti)", cell_format_catalogo)
worksheet.write(5, 7, "Primo Ingresso Normalizzato", cell_format_catalogo)
worksheet.write(5, 8, "Ultimo Ingresso Normalizzato", cell_format_catalogo)
worksheet.write(5, 9, "Tempo Reale Online", cell_format_catalogo)


row = 6
for x in sorted_csv:
    if x.tempo_online == 0:
        x.tempo_online = 1
    worksheet.write(row, 0, x.nome, cell_format_catalogo_nome)
    worksheet.write(row, 1, x.email, cell_format_catalogo_nome)
    worksheet.write(row, 2, x.join_date, cell_format_catalogo)
    worksheet.write(row, 3, x.orario_entrata, cell_format_catalogo)
    worksheet.write(row, 4, x.leave_date, cell_format_catalogo)
    worksheet.write(row, 5, x.orario_uscita, cell_format_catalogo)
    worksheet.write(row, 6, x.tempo_online, cell_format_catalogo)
    primo_ingresso = x.primo_ingresso_normalizzato
    ultimo_ingresso = x.ultimo_ingresso_normalizzato

    tmp_primo_ingresso = datetime.strptime(primo_ingresso, "%H:%M").time()
    tmp_ultimo_ingresso = datetime.strptime(ultimo_ingresso, "%H:%M").time()
    if self.OrarioEntrata.Value:
        primo_ingresso = datetime.strptime(primo_ingresso, "%H:%M").time()
        try:
            orario_entrata = datetime.strptime(self.OrarioEntrata.Value, "%H:%M").time()
        except:
            wx.MessageBox("Orario entrata sbagliato")

        if primo_ingresso < orario_entrata:
            primo_ingresso = orario_entrata
            tmp_primo_ingresso = primo_ingresso
            primo_ingresso = primo_ingresso.strftime("%H:%M")

        else:
            primo_ingresso = x.primo_ingresso_normalizzato

    if self.OrarioUscita.Value:
        ultimo_ingresso = datetime.strptime(ultimo_ingresso, "%H:%M").time()
        try:
            orario_uscita = datetime.strptime(self.OrarioUscita.Value, "%H:%M").time()
        except:
            wx.MessageBox("Orario uscita sbagliato")

        if ultimo_ingresso > orario_uscita:
            ultimo_ingresso = orario_uscita
            tmp_ultimo_ingresso = ultimo_ingresso
            ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
        elif ultimo_ingresso < orario_entrata:
            ultimo_ingresso = orario_entrata
            ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
        else:
            ultimo_ingresso = x.ultimo_ingresso_normalizzato
    if str(primo_ingresso) != str(ultimo_ingresso):
        x.tempo_online = self.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
    else:
        x.tempo_online = "0:00"


    worksheet.write(row, 7, primo_ingresso, cell_format_catalogo)
    worksheet.write(row, 8, ultimo_ingresso, cell_format_catalogo)
    worksheet.write(row, 9, x.tempo_online, cell_format_catalogo)
    #worksheet.write(row, 9, self.NormalizzaQuartoDora(x.tempo_online), cell_format_catalogo)
    row += 1
#DIVISORE------------------------------------------------------------------------------------------------------
worksheet = workbook.add_worksheet(name="ReportDiConsegna")

worksheet.set_row(0, 15)
worksheet.set_row(1, 70)
worksheet.set_row(2, 15)
worksheet.set_row(3, 35)
worksheet.set_row(4, 15)
worksheet.set_row(5, 50)
worksheet.set_row(6, 20)




worksheet.set_column(0, 0, 33.38)
worksheet.set_column(1, 1, 26.04)
worksheet.set_column(2, 2, 13.38)
worksheet.set_column(3, 3, 9.61)
worksheet.set_column(4, 4, 11.90)
worksheet.set_column(5, 5, 9.61)
worksheet.set_column(6, 6, 15.19)
worksheet.set_column(7, 7, 15.19)
worksheet.set_column(8, 8, 15.19)
worksheet.set_column(9, 9, 15.19)



worksheet.write(0, 0, "ID riunione", cell_format_ID_riunione)
worksheet.write(0, 1, "Argomento", cell_format_ID_riunione)
worksheet.write(0, 2, "Ora di inizio", cell_format_ID_riunione)
worksheet.write(0, 3, "", cell_format_ID_riunione)
worksheet.write(0, 4, "Ora di fine", cell_format_ID_riunione)
worksheet.write(0, 5, "", cell_format_ID_riunione)
worksheet.write(0, 6, "Email dell’utente", cell_format_ID_riunione)

worksheet.write(1, 0, header_generic_list[0].ID, cell_format_ID_riunione)
worksheet.write(1, 1, header_generic_list[0].Argomento, cell_format_ID_riunione_bold)
worksheet.write(1, 2, header_generic_list[0].Ora_inizio[0:10], cell_format_ID_riunione)
worksheet.write(1, 3, header_generic_list[0].Ora_inizio[10:], cell_format_ID_riunione)
worksheet.write(1, 4, header_generic_list[0].Ora_Fine[0:10], cell_format_ID_riunione)
worksheet.write(1, 5, header_generic_list[0].Ora_Fine[10:], cell_format_ID_riunione)
worksheet.write(1, 6, header_generic_list[0].Email, cell_format_ID_riunione)

tmp_primo_ingresso = datetime.strptime( self.OrarioEntrata.Value, "%H:%M").time()
tmp_ultimo_ingresso = datetime.strptime(self.OrarioUscita.Value, "%H:%M").time()
numero_ore = self.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)

minutes = int(numero_ore[2:])
minutes = minutes/60
hour = int(numero_ore[:1])
orario = (hour * 60) + (minutes * 60)
worksheet.write(2, 0, "Tutor", cell_format_ID_riunione)
worksheet.write(2, 1, self.NomeTutor.Value, cell_format_ID_riunione)
worksheet.write(2, 2, "Ora inizio", cell_format_ID_riunione)
worksheet.write(2, 3, self.OrarioEntrata.Value, cell_format_ID_riunione)
worksheet.write(2, 4, "N. ore", cell_format_ID_riunione)
worksheet.write(2, 5, numero_ore, cell_format_ID_riunione)
worksheet.write(2, 6, "", cell_format_ID_riunione)



worksheet.write(3, 0, "Docente", cell_format_ID_riunione)
worksheet.write(3, 1, self.NomeDocente.Value, cell_format_ID_riunione)
worksheet.write(3, 2, "Ora Fine", cell_format_ID_riunione)
worksheet.write(3, 3, self.OrarioUscita.Value, cell_format_ID_riunione)
worksheet.write(3, 4, "Durata in min reale", cell_format_ID_riunione)
worksheet.write(3, 5, orario, cell_format_ID_riunione)
worksheet.write(3, 6, "", cell_format_ID_riunione)

worksheet.write(4, 0, "", cell_format_catalogo)
worksheet.write(4, 1, "", cell_format_catalogo)
worksheet.write(4, 2, "", cell_format_catalogo)
worksheet.write(4, 3, "", cell_format_catalogo)
worksheet.write(4, 4, "", cell_format_catalogo)
worksheet.write(4, 5, "", cell_format_catalogo)
worksheet.write(4, 6, "", cell_format_catalogo)

worksheet.write(5, 0, "Nome (nome originale)", cell_format_catalogo)
worksheet.write(5, 1, "Email dell'utente", cell_format_catalogo)
worksheet.write(5, 2, "Join Date", cell_format_catalogo)
worksheet.write(5, 3, "Ora di Entrata", cell_format_catalogo)
worksheet.write(5, 4, "Leave Date", cell_format_catalogo)
worksheet.write(5, 5, "Ora di Uscita", cell_format_catalogo)
worksheet.write(5, 6, "Primo Ingresso Normalizzato", cell_format_catalogo)
worksheet.write(5, 7, "Ultimo Ingresso Normalizzato", cell_format_catalogo)
worksheet.write(5, 8, "Tempo Effettivo Online", cell_format_catalogo)
worksheet.write(5, 9, "Tempo Normalizzato al quarto d'ora", cell_format_catalogo)

row = 6
tmp.sort(key=lambda x: x.email)
for x in tmp:
    worksheet.write(row, 0, x.nome, cell_format_catalogo_nome)
    worksheet.write(row, 1, x.email, cell_format_catalogo_nome)
    worksheet.write(row, 2, x.join_date, cell_format_catalogo)
    worksheet.write(row, 3, x.orario_entrata, cell_format_catalogo)
    worksheet.write(row, 4, x.leave_date, cell_format_catalogo)
    worksheet.write(row, 5, x.orario_uscita, cell_format_catalogo)
    primo_ingresso = x.primo_ingresso_normalizzato
    ultimo_ingresso = x.ultimo_ingresso_normalizzato

    tmp_primo_ingresso = datetime.strptime(primo_ingresso, "%H:%M").time()
    tmp_ultimo_ingresso = datetime.strptime(ultimo_ingresso, "%H:%M").time()
    if self.OrarioEntrata.Value:
        primo_ingresso = datetime.strptime(primo_ingresso, "%H:%M").time()
        try:
            orario_entrata = datetime.strptime(self.OrarioEntrata.Value, "%H:%M").time()
        except:
            wx.MessageBox("Orario entrata sbagliato")

        if primo_ingresso < orario_entrata:
            primo_ingresso = orario_entrata
            tmp_primo_ingresso = primo_ingresso
            primo_ingresso = primo_ingresso.strftime("%H:%M")

        else:
            primo_ingresso = x.primo_ingresso_normalizzato

    if self.OrarioUscita.Value:
        ultimo_ingresso = datetime.strptime(ultimo_ingresso, "%H:%M").time()
        try:
            orario_uscita = datetime.strptime(self.OrarioUscita.Value, "%H:%M").time()
        except:
            wx.MessageBox("Orario uscita sbagliato")

        if ultimo_ingresso > orario_uscita:
            ultimo_ingresso = orario_uscita
            tmp_ultimo_ingresso = ultimo_ingresso
            ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
        elif ultimo_ingresso < orario_entrata:
            ultimo_ingresso = orario_entrata
            ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
        else:
            ultimo_ingresso = x.ultimo_ingresso_normalizzato
    if str(primo_ingresso) != str(ultimo_ingresso):
        x.tempo_online = self.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
    else:
        x.tempo_online = "0:00"


    worksheet.write(row, 6, primo_ingresso, cell_format_catalogo)
    worksheet.write(row, 7, ultimo_ingresso, cell_format_catalogo)
    worksheet.write(row, 8, x.tempo_online, cell_format_catalogo)
    worksheet.write(row, 9, self.NormalizzaQuartoDora(x.tempo_online), cell_format_catalogo)
    row += 1
row += 1
worksheet.insert_image(row, 5, "immagine.png", {"x_scale": 0.25, "y_scale": 0.25, "x_offset": 50, "y_offset": 10})
workbook.close()
if self.DirPicker.GetPath():
    os.startfile(f"{path}/{header_generic_list[0].ID}.xlsx")
else:
    os.startfile(f"{header_generic_list[0].ID}.xlsx")

def CalcolaTempoOnline(self, primo_ingresso, ultimo_ingresso):

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
    print(tempo_finale)
    return tempo_finale

def NormalizzaQuartoDora(self, tempo_online):
    print(tempo_online)
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