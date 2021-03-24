import xlsxwriter
from Controllers.BodyCsv import BodyCsv as CsvFunc
from datetime import datetime
def WriteAttendeeReportHeader(HeaderCsv, file_name, primo_ingresso, ultimo_ingresso, tutor, docente, nome_sheet, workbook = None):

    if workbook == None:
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet(name=nome_sheet)
    else:
        worksheet = workbook.add_worksheet(name=nome_sheet)

    cell_format_ID_riunione = workbook.add_format({"border": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_ID_riunione_bold = workbook.add_format({"border": True, "bold": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo = workbook.add_format({"font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo_nome = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left"})
    cell_format_catalogo_nome_bold = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "bold": True})
    cell_format_catalogo_nome_border = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "top": 1, "bottom": 1,"valign": "vcenter", "text_wrap": True, "align": "left"})
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


    worksheet.write(1, 0, HeaderCsv.ID, cell_format_ID_riunione)
    worksheet.write(1, 1, HeaderCsv.Argomento, cell_format_ID_riunione_bold)
    worksheet.write(1, 2, HeaderCsv.OraInizio[0:10], cell_format_ID_riunione)
    worksheet.write(1, 3, HeaderCsv.OraInizio[10:], cell_format_ID_riunione)
    worksheet.write(1, 4, HeaderCsv.OraFine[0:10], cell_format_ID_riunione)
    worksheet.write(1, 5, HeaderCsv.OraFine[10:], cell_format_ID_riunione)
    worksheet.write(1, 6, HeaderCsv.Email, cell_format_ID_riunione)
    print(primo_ingresso)
    tmp_primo_ingresso = datetime.strptime(primo_ingresso, "%H:%M").time()
    tmp_ultimo_ingresso = datetime.strptime(ultimo_ingresso, "%H:%M").time()
    numero_ore = CsvFunc.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
    numero_ore = numero_ore.split(":")
    minutes = int(numero_ore[1])
    minutes = minutes/60
    hour = int(numero_ore[0])
    orario = (hour * 60) + (minutes * 60)
    worksheet.write(2, 0, "Tutor", cell_format_ID_riunione)
    worksheet.write(2, 1, tutor, cell_format_ID_riunione_bold)
    worksheet.write(2, 2, "Ora inizio", cell_format_ID_riunione)
    worksheet.write(2, 3, primo_ingresso, cell_format_ID_riunione)
    worksheet.write(2, 4, "N. ore", cell_format_ID_riunione)
    worksheet.write(2, 5, numero_ore[0]+":"+numero_ore[1], cell_format_ID_riunione)
    worksheet.write(2, 6, "", cell_format_ID_riunione)


    worksheet.write(3, 0, "Docente", cell_format_ID_riunione)
    worksheet.write(3, 1, docente, cell_format_ID_riunione_bold)
    worksheet.write(3, 2, "Ora Fine", cell_format_ID_riunione)
    worksheet.write(3, 3, ultimo_ingresso, cell_format_ID_riunione)
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


    return workbook
    row = 6


def WriteAttendeeReportBody(workbook, BodyCsv, ora_entrata, ora_uscita):
    worksheet = workbook.get_worksheet_by_name("AttandeeReport")
    cell_format_ID_riunione = workbook.add_format({"border": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_ID_riunione_bold = workbook.add_format({"border": True, "bold": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo = workbook.add_format({"font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo_bg = workbook.add_format({"bg_color": "d9d9d9","font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo_nome = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left"})
    cell_format_catalogo_nome_bg = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "bg_color": "d9d9d9"})
    cell_format_catalogo_nome_border = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "top": 1, "bottom": 1, "valign": "vcenter", "text_wrap": True, "align": "left"})



    worksheet.write(5, 0, "Nome (nome originale)", cell_format_catalogo_nome_border)
    worksheet.write(5, 1, "Email dell'utente", cell_format_catalogo_nome_border)
    worksheet.write(5, 2, "Join Date", cell_format_catalogo_nome_border)
    worksheet.write(5, 3, "Ora di Entrata", cell_format_catalogo_nome_border)
    worksheet.write(5, 4, "Leave Date", cell_format_catalogo_nome_border)
    worksheet.write(5, 5, "Ora di Uscita", cell_format_catalogo_nome_border)
    worksheet.write(5, 6, "Durata (Minuti)", cell_format_catalogo_nome_border)
    worksheet.write(5, 7, "Primo Ingresso Normalizzato", cell_format_catalogo_nome_border)
    worksheet.write(5, 8, "Ultimo Ingresso Normalizzato", cell_format_catalogo_nome_border)
    worksheet.write(5, 9, "Tempo Reale Online", cell_format_catalogo_nome_border)

    row = 6
    flg = 0
    for y in BodyCsv:
        for x in y:
            if flg == 0:
                one_format = cell_format_catalogo_nome
                two_format = cell_format_catalogo_nome
                flg = 1
            else:
                one_format = cell_format_catalogo_nome
                two_format = cell_format_catalogo
                flg = 0
            worksheet.write(row, 0, x.Nome, one_format)
            worksheet.write(row, 1, x.Email, one_format)
            worksheet.write(row, 2, x.JoinDate.strftime("%d/%m/%Y"), two_format)
            worksheet.write(row, 3, x.OraEntrata.strftime("%H:%M:%S"), two_format)
            worksheet.write(row, 4, x.LeaveDate.strftime("%d/%m/%Y"), two_format)
            worksheet.write(row, 5, x.OrarioUscita.strftime("%H:%M:%S"), two_format)
            worksheet.write(row, 6, x.TempoOnline, two_format)

            tmp_primo_ingresso = x.IngressoNormalizzato
            tmp_ultimo_ingresso = x.UscitaNormalizzato
            if ora_entrata:
                primo_ingresso = x.IngressoNormalizzato
                try:
                    orario_entrata = datetime.strptime(ora_entrata, "%H:%M").time()
                except:
                    print("ora entrata sbagliata")

                if primo_ingresso < orario_entrata:
                    primo_ingresso = orario_entrata
                    tmp_primo_ingresso = primo_ingresso
                    primo_ingresso = primo_ingresso.strftime("%H:%M")
                else:
                    primo_ingresso = x.IngressoNormalizzato.strftime("%H:%M")

            if ora_uscita:
                ultimo_ingresso = x.UscitaNormalizzato
                try:
                    orario_uscita = datetime.strptime(ora_uscita, "%H:%M").time()
                except:
                    print("orario uscita sbagliato")

                if ultimo_ingresso > orario_uscita:
                    ultimo_ingresso = orario_uscita
                    tmp_ultimo_ingresso = ultimo_ingresso
                    ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
                elif ultimo_ingresso < orario_entrata:
                    ultimo_ingresso = orario_entrata
                    ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
                else:
                    ultimo_ingresso = x.UscitaNormalizzato.strftime("%H:%M")
            if str(primo_ingresso) != str(ultimo_ingresso):
                tempo_online = CsvFunc.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
            else:
                tempo_online = "0:00"


            worksheet.write(row, 7, primo_ingresso, two_format)
            worksheet.write(row, 8, ultimo_ingresso, two_format)
            worksheet.write(row, 9, tempo_online, two_format)
            #worksheet.write(row, 9, self.NormalizzaQuartoDora(x.tempo_online), cell_format_catalogo)
            row += 1
        cell_format_catalogo_nome_border = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "top": 1, "valign": "vcenter", "text_wrap": True, "align": "left"})
    worksheet.write(row, 9, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 8, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 7, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 6, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 5, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 4, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 3, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 2, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 1, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 0, "", cell_format_catalogo_nome_border)
    row += 2
    #worksheet.insert_image(row, 5, "immagine.png", {"x_scale": 0.25, "y_scale": 0.25, "x_offset": 50, "y_offset": 10})
    return workbook
        #if self.DirPicker.GetPath():
        #    os.startfile(f"{path}/{header_generic_list[0].ID}.xlsx")
        #else:
        #    os.startfile(f"{header_generic_list[0].ID}.xlsx")


def WriteReportDiConsegnaBody(workbook, BodyCsv, ora_entrata, ora_uscita):
    worksheet = workbook.get_worksheet_by_name("ReportDiConsegna")
    cell_format_ID_riunione = workbook.add_format({"border": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_ID_riunione_bold = workbook.add_format({"border": True, "bold": True, "font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo = workbook.add_format({"font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo_bg = workbook.add_format({"bg_color": "d9d9d9","font_name": "Calibri", "font_size": 11, "valign": "vcenter", "text_wrap": True, "align": "left"})
    cell_format_catalogo_nome = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left"})
    cell_format_catalogo_nome_bg = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "bg_color": "d9d9d9"})
    cell_format_catalogo_nome_border = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "top": 1, "bottom": 1, "valign": "vcenter", "text_wrap": True, "align": "left"})


    worksheet.write(5, 0, "Nome (nome originale)", cell_format_catalogo_nome_border)
    worksheet.write(5, 1, "Email dell'utente", cell_format_catalogo_nome_border)
    worksheet.write(5, 2, "Join Date", cell_format_catalogo_nome_border)
    worksheet.write(5, 3, "Ora di Entrata", cell_format_catalogo_nome_border)
    worksheet.write(5, 4, "Leave Date", cell_format_catalogo_nome_border)
    worksheet.write(5, 5, "Ora di Uscita", cell_format_catalogo_nome_border)
    worksheet.write(5, 6, "Primo Ingresso Normalizzato", cell_format_catalogo_nome_border)
    worksheet.write(5, 7, "Ultimo Ingresso Normalizzato", cell_format_catalogo_nome_border)
    worksheet.write(5, 8, "Tempo Effettivo Online", cell_format_catalogo_nome_border)
    worksheet.write(5, 9, "Tempo Normalizzato al quarto d'ora", cell_format_catalogo_nome_border)

    row = 6
    flg = 0
    for x in BodyCsv:

        if flg == 0:
            one_format = cell_format_catalogo_nome_bg
            two_format = cell_format_catalogo_bg
            flg = 1
        else:
            one_format = cell_format_catalogo_nome
            two_format = cell_format_catalogo
            flg = 0
        worksheet.write(row, 0, x.Nome, one_format)
        worksheet.write(row, 1, x.Email, one_format)
        worksheet.write(row, 2, x.JoinDate.strftime("%d/%m/%Y"), two_format)
        worksheet.write(row, 3, x.OraEntrata.strftime("%H:%M:%S"), two_format)
        worksheet.write(row, 4, x.LeaveDate.strftime("%d/%m/%Y"), two_format)
        worksheet.write(row, 5, x.OrarioUscita.strftime("%H:%M:%S"), two_format)
        #worksheet.write(row, 6, x.TempoOnline, two_format)

        tmp_primo_ingresso = x.IngressoNormalizzato
        tmp_ultimo_ingresso = x.UscitaNormalizzato
        if ora_entrata:
            primo_ingresso = x.IngressoNormalizzato
            try:
                orario_entrata = datetime.strptime(ora_entrata, "%H:%M").time()
            except:
                print("ora entrata sbagliata")

            if primo_ingresso < orario_entrata:
                primo_ingresso = orario_entrata
                tmp_primo_ingresso = primo_ingresso
                primo_ingresso = primo_ingresso.strftime("%H:%M")
            else:
                primo_ingresso = x.IngressoNormalizzato.strftime("%H:%M")

        if ora_uscita:
            ultimo_ingresso = x.UscitaNormalizzato
            try:
                orario_uscita = datetime.strptime(ora_uscita, "%H:%M").time()
            except:
                print("orario uscita sbagliato")

            if ultimo_ingresso > orario_uscita:
                ultimo_ingresso = orario_uscita
                tmp_ultimo_ingresso = ultimo_ingresso
                ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
            elif ultimo_ingresso < orario_entrata:
                ultimo_ingresso = orario_entrata
                ultimo_ingresso = ultimo_ingresso.strftime("%H:%M")
            else:
                ultimo_ingresso = x.UscitaNormalizzato.strftime("%H:%M")
        if str(primo_ingresso) != str(ultimo_ingresso):
            tempo_online = CsvFunc.CalcolaTempoOnline(tmp_primo_ingresso, tmp_ultimo_ingresso)
        else:
            tempo_online = "0:00"


        worksheet.write(row, 6, primo_ingresso, two_format)
        worksheet.write(row, 7, ultimo_ingresso, two_format)
        worksheet.write(row, 8, tempo_online, two_format)
        worksheet.write(row, 9, CsvFunc.NormalizzaQuartoDora(tempo_online), two_format)
        #worksheet.write(row, 9, self.NormalizzaQuartoDora(x.tempo_online), cell_format_catalogo)
        row += 1
    cell_format_catalogo_nome_border = workbook.add_format({"font_name": "Calibri", "font_size": 11, "align": "left", "top": 1, "valign": "vcenter", "text_wrap": True, "align": "left"})
    worksheet.write(row, 9, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 8, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 7, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 6, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 5, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 4, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 3, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 2, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 1, "", cell_format_catalogo_nome_border)
    worksheet.write(row, 0, "", cell_format_catalogo_nome_border)
    row += 2
    worksheet.insert_image(row, 5, "immagine.png", {"x_scale": 0.25, "y_scale": 0.25, "x_offset": 50, "y_offset": 10})
    return workbook