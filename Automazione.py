# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Oct 26 2018)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import wx.aui
import wx.grid

###########################################################################
## Class Mainframe
###########################################################################

class Mainframe ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Francesco De Angelis - Automatron Automazione Reportistica ", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.MAXIMIZE|wx.MAXIMIZE_BOX|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
		self.m_mgr = wx.aui.AuiManager()
		self.m_mgr.SetManagedWindow( self )
		self.m_mgr.SetFlags(wx.aui.AUI_MGR_DEFAULT)

		self.MainFrameToolBar = wx.MenuBar( 0 )
		self.ReportMenu = wx.Menu()
		self.TrasformatoreExcelMenu = wx.MenuItem( self.ReportMenu, wx.ID_ANY, u"Trasformatore Excel", wx.EmptyString, wx.ITEM_NORMAL )
		self.ReportMenu.Append( self.TrasformatoreExcelMenu )

		self.MainFrameToolBar.Append( self.ReportMenu, u"Report" )

		self.SetMenuBar( self.MainFrameToolBar )

		self.MainFrameNotebook = wx.aui.AuiNotebook( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( 1920,1080 ), wx.aui.AUI_NB_DEFAULT_STYLE )
		self.m_mgr.AddPane( self.MainFrameNotebook, wx.aui.AuiPaneInfo() .Left() .PinButton( True ).Dock().Resizable().FloatingSize( wx.DefaultSize ).CentrePane() )



		self.m_mgr.Update()
		self.Centre( wx.BOTH )

		# Connect Events
		self.Bind( wx.EVT_MENU, self.OnMenuTrasformatore, id = self.TrasformatoreExcelMenu.GetId() )

	def __del__( self ):
		self.m_mgr.UnInit()



	# Virtual event handlers, overide them in your derived class
	def OnMenuTrasformatore( self, event ):
		event.Skip()


###########################################################################
## Class TrasformatoreExcelPanel
###########################################################################

class TrasformatoreExcelPanel ( wx.Panel ):

	def __init__( self, parent, id = wx.ID_ANY, pos = wx.DefaultPosition, size = wx.Size( -1,-1 ), style = wx.TAB_TRAVERSAL, name = wx.EmptyString ):
		wx.Panel.__init__ ( self, parent, id = id, pos = pos, size = size, style = style, name = name )

		bTrasformatoreExcel = wx.BoxSizer( wx.VERTICAL )

		bSelector = wx.BoxSizer( wx.HORIZONTAL )

		self.CsvFilePickerLabel = wx.StaticText( self, wx.ID_ANY, u"Csv File:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.CsvFilePickerLabel.Wrap( -1 )

		self.CsvFilePickerLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bSelector.Add( self.CsvFilePickerLabel, 1, wx.ALL, 20 )

		self.CsvFilePicker = wx.FilePickerCtrl( self, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE )
		self.CsvFilePicker.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bSelector.Add( self.CsvFilePicker, 5, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bSelector, 1, wx.EXPAND, 5 )

		bOutput = wx.BoxSizer( wx.HORIZONTAL )

		self.OutputDirectoryLabel = wx.StaticText( self, wx.ID_ANY, u"Output Directory:", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OutputDirectoryLabel.Wrap( -1 )

		self.OutputDirectoryLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOutput.Add( self.OutputDirectoryLabel, 1, wx.ALL, 20 )

		self.DirPicker = wx.DirPickerCtrl( self, wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition, wx.DefaultSize, wx.DIRP_DEFAULT_STYLE )
		self.DirPicker.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOutput.Add( self.DirPicker, 5, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bOutput, 1, wx.EXPAND, 5 )

		bTutor = wx.BoxSizer( wx.HORIZONTAL )

		self.TutorLabel = wx.StaticText( self, wx.ID_ANY, u"Tutor", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.TutorLabel.Wrap( -1 )

		self.TutorLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bTutor.Add( self.TutorLabel, 1, wx.ALL, 20 )

		NomeTutorChoices = []
		self.NomeTutor = wx.ComboBox( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, NomeTutorChoices, 0 )
		self.NomeTutor.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bTutor.Add( self.NomeTutor, 5, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bTutor, 1, wx.EXPAND, 5 )

		bDocente = wx.BoxSizer( wx.HORIZONTAL )

		self.DocenteLabel = wx.StaticText( self, wx.ID_ANY, u"Docente", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.DocenteLabel.Wrap( -1 )

		self.DocenteLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bDocente.Add( self.DocenteLabel, 1, wx.ALL, 20 )

		NomeDocenteChoices = []
		self.NomeDocente = wx.ComboBox( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, NomeDocenteChoices, 0 )
		self.NomeDocente.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bDocente.Add( self.NomeDocente, 5, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bDocente, 1, wx.EXPAND, 5 )

		bOrario = wx.BoxSizer( wx.HORIZONTAL )

		self.OrarioEntrataLabel = wx.StaticText( self, wx.ID_ANY, u"Orario Entrata", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OrarioEntrataLabel.Wrap( -1 )

		self.OrarioEntrataLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOrario.Add( self.OrarioEntrataLabel, 1, wx.ALL, 20 )

		self.OrarioEntrata = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OrarioEntrata.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOrario.Add( self.OrarioEntrata, 2, wx.ALL, 20 )

		self.OrarioUscitaLabel = wx.StaticText( self, wx.ID_ANY, u"Orario Uscita", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OrarioUscitaLabel.Wrap( -1 )

		self.OrarioUscitaLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOrario.Add( self.OrarioUscitaLabel, 1, wx.ALL, 20 )

		self.OrarioUscita = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		self.OrarioUscita.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bOrario.Add( self.OrarioUscita, 2, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bOrario, 1, wx.EXPAND, 5 )

		bTable = wx.BoxSizer( wx.HORIZONTAL )

		self.TabellaTutor = wx.grid.Grid( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( -1,260 ), 0 )

		# Grid
		self.TabellaTutor.CreateGrid( 1, 1 )
		self.TabellaTutor.EnableEditing( True )
		self.TabellaTutor.EnableGridLines( True )
		self.TabellaTutor.EnableDragGridSize( False )
		self.TabellaTutor.SetMargins( 0, 0 )

		# Columns
		self.TabellaTutor.SetColSize( 0, 116 )
		self.TabellaTutor.AutoSizeColumns()
		self.TabellaTutor.EnableDragColMove( True )
		self.TabellaTutor.EnableDragColSize( True )
		self.TabellaTutor.SetColLabelSize( 30 )
		self.TabellaTutor.SetColLabelValue( 0, u"Nome Tutor" )
		self.TabellaTutor.SetColLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

		# Rows
		self.TabellaTutor.EnableDragRowSize( True )
		self.TabellaTutor.SetRowLabelSize( 80 )
		self.TabellaTutor.SetRowLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

		# Label Appearance
		self.TabellaTutor.SetLabelFont( wx.Font( 14, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		# Cell Defaults
		self.TabellaTutor.SetDefaultCellFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )
		self.TabellaTutor.SetDefaultCellAlignment( wx.ALIGN_LEFT, wx.ALIGN_TOP )
		bTable.Add( self.TabellaTutor, 1, wx.ALL, 20 )

		self.TabellaDocente = wx.grid.Grid( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( -1,260 ), 0 )

		# Grid
		self.TabellaDocente.CreateGrid( 1, 1 )
		self.TabellaDocente.EnableEditing( True )
		self.TabellaDocente.EnableGridLines( True )
		self.TabellaDocente.EnableDragGridSize( False )
		self.TabellaDocente.SetMargins( 0, 0 )

		# Columns
		self.TabellaDocente.AutoSizeColumns()
		self.TabellaDocente.EnableDragColMove( True )
		self.TabellaDocente.EnableDragColSize( True )
		self.TabellaDocente.SetColLabelSize( 30 )
		self.TabellaDocente.SetColLabelValue( 0, u"Nome Docente" )
		self.TabellaDocente.SetColLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

		# Rows
		self.TabellaDocente.EnableDragRowSize( True )
		self.TabellaDocente.SetRowLabelSize( 80 )
		self.TabellaDocente.SetRowLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

		# Label Appearance
		self.TabellaDocente.SetLabelFont( wx.Font( 14, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		# Cell Defaults
		self.TabellaDocente.SetDefaultCellFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )
		self.TabellaDocente.SetDefaultCellAlignment( wx.ALIGN_LEFT, wx.ALIGN_TOP )
		bTable.Add( self.TabellaDocente, 1, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bTable, 4, wx.EXPAND, 5 )

		bAddTutor = wx.BoxSizer( wx.HORIZONTAL )

		self.AggiungiTutorLabel = wx.StaticText( self, wx.ID_ANY, u"Aggiungi Tutor", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.AggiungiTutorLabel.Wrap( -1 )

		self.AggiungiTutorLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bAddTutor.Add( self.AggiungiTutorLabel, 1, wx.ALL, 20 )

		self.AggiungiTutor = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor.Add( self.AggiungiTutor, 4, wx.ALL, 20 )

		self.AggiungiTutorButton = wx.Button( self, wx.ID_ANY, u"Aggiungi", wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor.Add( self.AggiungiTutorButton, 1, wx.ALL, 20 )

		self.RimuoviTutorButton = wx.Button( self, wx.ID_ANY, u"Rimuovi", wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor.Add( self.RimuoviTutorButton, 1, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bAddTutor, 1, wx.EXPAND, 5 )

		bAddTutor1 = wx.BoxSizer( wx.HORIZONTAL )

		self.AggiungiDocenteLabel = wx.StaticText( self, wx.ID_ANY, u"Aggiungi Docente", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.AggiungiDocenteLabel.Wrap( -1 )

		self.AggiungiDocenteLabel.SetFont( wx.Font( 11, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bAddTutor1.Add( self.AggiungiDocenteLabel, 1, wx.ALL, 20 )

		self.AggiungiDocente = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor1.Add( self.AggiungiDocente, 4, wx.ALL, 20 )

		self.AggiungiDocenteButton = wx.Button( self, wx.ID_ANY, u"Aggiungi", wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor1.Add( self.AggiungiDocenteButton, 1, wx.ALL, 20 )

		self.RimuoviDocenteButton = wx.Button( self, wx.ID_ANY, u"Rimuovi", wx.DefaultPosition, wx.DefaultSize, 0 )
		bAddTutor1.Add( self.RimuoviDocenteButton, 1, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bAddTutor1, 1, wx.EXPAND, 5 )

		bStampa = wx.BoxSizer( wx.HORIZONTAL )

		self.StampaExcel = wx.Button( self, wx.ID_ANY, u"Stampa Excel", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.StampaExcel.SetFont( wx.Font( 14, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bStampa.Add( self.StampaExcel, 1, wx.ALL, 20 )

		self.StampaPdf = wx.Button( self, wx.ID_ANY, u"Stampa PDF", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.StampaPdf.SetFont( wx.Font( 14, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "Verdana" ) )

		bStampa.Add( self.StampaPdf, 1, wx.ALL, 20 )


		bTrasformatoreExcel.Add( bStampa, 1, wx.EXPAND, 5 )


		self.SetSizer( bTrasformatoreExcel )
		self.Layout()
		bTrasformatoreExcel.Fit( self )

		# Connect Events
		self.TabellaTutor.Bind( wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.OnTutorRightClick )
		self.TabellaDocente.Bind( wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.OnDocenteRightClick )
		self.AggiungiTutorButton.Bind( wx.EVT_BUTTON, self.AggiungiTutorEvent )
		self.RimuoviTutorButton.Bind( wx.EVT_BUTTON, self.RimuoviTutorEvent )
		self.AggiungiDocenteButton.Bind( wx.EVT_BUTTON, self.AggiungiDocenteEvent )
		self.RimuoviDocenteButton.Bind( wx.EVT_BUTTON, self.RimuoviDocenteEvent )
		self.StampaExcel.Bind( wx.EVT_BUTTON, self.StampaExcelEvent )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def OnTutorRightClick( self, event ):
		event.Skip()

	def OnDocenteRightClick( self, event ):
		event.Skip()

	def AggiungiTutorEvent( self, event ):
		event.Skip()

	def RimuoviTutorEvent( self, event ):
		event.Skip()

	def AggiungiDocenteEvent( self, event ):
		event.Skip()

	def RimuoviDocenteEvent( self, event ):
		event.Skip()

	def StampaExcelEvent( self, event ):
		event.Skip()


