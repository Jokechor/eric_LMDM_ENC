import sys, win32com.client
import os
import pymssql
import datetime
import logging
import OpOutlook


class SAP:
    server = "147.128.98.27"
    database = "ESS_NJ_LMDM_MD"
    user = ""
    password = ""
    #make security info invisible

    def __init__(self, SAPSystem):
        if os.path.exists(r"C:\Python 3.7.1\SAPClassLog.txt"):
            os.remove(r"C:\Python 3.7.1\SAPClassLog.txt")
        self.SAPConn = False
        self.session = None
        self.SAPSystem = SAPSystem
        self.sqlList = []  # Store the SQL sentence
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        self.logfile = os.path.join(os.path.abspath(".."), "SAPClassLog.txt")
        self.logFormat = logging.Formatter("%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s")
        self.loghandler = logging.FileHandler(self.logfile, mode="a")
        self.loghandler.setLevel(logging.INFO)
        self.loghandler.setFormatter(self.logFormat)
        self.logger.addHandler(self.loghandler)
        self.logger.info("MAIN Exe starts")

    def ConnSAP(self):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.Children(0)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        SystemOpenCount = application.Children.count  # To confirm whether log on the target system
        if SystemOpenCount != 0:
            for i in range(SystemOpenCount):
                if self.SAPSystem in application.Children(i).Children(0).passportsystemid:
                    self.SAPConn = True
        if self.SAPConn == False:
            self.logger.warning("SAP is not log on to " + self.SAPSystem)
            print("SAP is not log on to " + self.SAPSystem)
            return
        self.session = session

    def ExportMARA(self):
        if self.SAPConn is True:
            self.logger.info("Export MARA report starts")
            print("Export MARA report starts at " + str(datetime.datetime.now()))
            session = self.session
        else:
            self.logger.warning("SAP disconnected when exporting data")
            print("SAP disconnected when exporting data")
            return
        self.logger.info("Export MARA starts")
        print("Export MARA start at " + str(datetime.datetime.now()))
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZMARA"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").text = "2491"
        session.findById("wnd[0]/usr/ctxtSP$00002-LOW").caretPosition = 4
        session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press()
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]"
        ).text = "3006"
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).text = "3805"
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).setFocus()
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtSP$00001-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSP$00001-LOW").caretPosition = 1
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).currentCellColumn = "TEXT"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).selectedRows = "0"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\\New QC\\PDM"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MARA.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        self.logger.info("Export MARA Session ends")

    def ExportMARC(self):
        if self.SAPConn is True:
            self.logger.info("Export MARC starts")
            print("Export MARC report starts at " + str(datetime.datetime.now()))
            session = self.session
        else:
            self.logger.warning("SAP disconnected when exporting data")
            print("SAP disconnected when exporting data")
            return
        self.logger.info("Export MARC Session starts")
        print("Export MARC Session starts at " + str(datetime.datetime.now()))
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZE16N"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtS_TABLE-LOW").text = "MARC"
        session.findById("wnd[0]/usr/ctxtS_TABLE-LOW").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtI1-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtI2-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtI2-LOW").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_I2_%_APP_%-VALU_PUSH").press()
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]"
        ).text = "2491"
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]"
        ).text = "3006"
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).text = "3805"
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).setFocus()
        session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]"
        ).caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = "X"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).currentCellColumn = "TEXT"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).selectedRows = "0"
        session.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
        ).clickCurrentCell()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\\New QC\\PDM"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MARC.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        self.logger.info("Export MARC Session ends")

    def uploadMARAMRAC(self):
        MARA_filePath = r"C:\New QC\PDM\MARA.txt"
        MARC_filePath = r"C:\New QC\PDM\MARC.txt"
        origiinal_MARA_data = []  # List to store the orginal MARA data
        origiinal_MARC_data = []  # List to store the orginal MRAC data
        self.logger.info("Operate MARA text starts")
        print("Operate MARA text start at " + str(datetime.datetime.now()))
        with open(MARA_filePath, "r") as MARA_f:
            for MARA_line in MARA_f:
                if MARA_line != "\n" and "\t" in MARA_line:  # delete Error MARA_line
                    MARA_line = MARA_line.split("\t")  # Remove Tab text
                    if MARA_line[0] == "":
                        if MARA_line[1] != "Material":
                            # print(MARA_line)
                            for singalChar in MARA_line:
                                singalIndex = MARA_line.index(singalChar)
                                singalChar = singalChar.strip()
                                singalChar = singalChar.replace("\n", "")
                                MARA_line[singalIndex] = singalChar
                            origiinal_MARA_data.append(MARA_line)
                    # print(origiinal_MARA_data[0])

        with open(MARC_filePath, "r") as MARC_f:
            self.logger.info("Operate MARC text starts")
            print("Operate MARC text start at " + str(datetime.datetime.now()))
            for MARC_line in MARC_f:
                if MARC_line != "\n" and "\t" in MARC_line:  # delete Error MARC_line
                    MARC_line = MARC_line.split("\t")  # Remove Tab text
                    if MARC_line[0] == "":
                        for singalChar in MARC_line:
                            singalIndex = MARC_line.index(singalChar)
                            singalChar = singalChar.strip()  # Remove the Tab or Space char in singal cell
                            MARC_line[singalIndex] = singalChar
                        origiinal_MARC_data.append(MARC_line)
                    # print(origiinal_MARC_data[0])
        self.logger.info("Execute Initialization Procedure starts")
        print("Execute Initialization Procedure at " + str(datetime.datetime.now()))
        self.runSQL([], "", "INITIAL_TABLES")  # Clean up original temp tables
        self.logger.info("Upload MARA data starts")
        print("Upload MARA data start at " + str(datetime.datetime.now()))
        for i in range(len(origiinal_MARA_data)):
            while len(origiinal_MARA_data[i]) < 7:
                self.logger.info("MARA SQL Adding blank")
                origiinal_MARA_data[i].append("")
                # print(origiinal_MARA_data[i])
            MATNR = origiinal_MARA_data[i][1]
            MTART = origiinal_MARA_data[i][3]
            MATKL = origiinal_MARA_data[i][4]
            MSTAE = origiinal_MARA_data[i][5]
            ZZMVGR1 = origiinal_MARA_data[i][6]
            sql = (
                r"INSERT INTO [ESS_NJ_LMDM_MD].[dbo].[TEMPMARA] (MATNR, MTART, MATKL, MSTAE, ZZMVGR1) VALUES('"
                + MATNR
                + r"','"
                + MTART
                + r"', '"
                + MATKL
                + r"', '"
                + MSTAE
                + r"', '"
                + ZZMVGR1
                + r"')"
            )
            self.logger.info("MARA Insert SQL is " + sql)
            self.sqlList.append(sql)
        self.runSQL(self.sqlList)
        self.sqlList.clear()
        self.logger.info("Upload MARC data starts")
        print("Upload MARC data start at " + str(datetime.datetime.now()))
        for i in range(len(origiinal_MARC_data)):
            # print("Serial:" + str(i))
            # print(origiinal_MARC_data[i])
            while len(origiinal_MARC_data[i]) < len(origiinal_MARC_data[1]):
                self.logger.info("MARC SQL Adding blank")
                print("Adding...")
                origiinal_MARC_data[i].append("")
                # print(origiinal_MARC_data[i])
            if i != 0:
                # print(origiinal_MARC_data[i])
                MATNR = str(
                    origiinal_MARC_data[i][1]
                )  # When python get Cell's value, the value will keep its original format, e.g. folat will be number, text will be string
                WERKS = str(origiinal_MARC_data[i][2])
                LVORM = str(origiinal_MARC_data[i][3])
                MMSTA = str(origiinal_MARC_data[i][4])
                MMSTD = str(origiinal_MARC_data[i][5])
                MAABC = str(origiinal_MARC_data[i][6])
                EKGRP = str(origiinal_MARC_data[i][7])
                AUSME = str(origiinal_MARC_data[i][8])
                DISPR = str(origiinal_MARC_data[i][9])
                DISMM = str(origiinal_MARC_data[i][10])
                DISPO = str(origiinal_MARC_data[i][11])
                PLIFZ = str(origiinal_MARC_data[i][12])
                WEBAZ = str(origiinal_MARC_data[i][13])
                PERKZ = str(origiinal_MARC_data[i][14])
                DISLS = str(origiinal_MARC_data[i][15])
                BESKZ = str(origiinal_MARC_data[i][16])
                SOBSL = str(origiinal_MARC_data[i][17])
                MINBE = str(origiinal_MARC_data[i][18])
                EISBE = str(origiinal_MARC_data[i][19])
                BSTMI = str(origiinal_MARC_data[i][20])
                BSTMA = str(origiinal_MARC_data[i][21])
                BSTFE = str(origiinal_MARC_data[i][22])
                BSTRF = str(origiinal_MARC_data[i][23])
                MABST = str(origiinal_MARC_data[i][24])
                SBDKZ = str(origiinal_MARC_data[i][25])
                ALTSL = str(origiinal_MARC_data[i][26])
                MISKZ = str(origiinal_MARC_data[i][27])
                FHORI = str(origiinal_MARC_data[i][28])
                RGEKZ = str(origiinal_MARC_data[i][29])
                FEVOR = str(origiinal_MARC_data[i][30])
                BEARZ = str(origiinal_MARC_data[i][31])
                TRANZ = str(origiinal_MARC_data[i][32])
                DZEIT = str(origiinal_MARC_data[i][33])
                USEQU = str(origiinal_MARC_data[i][34])
                MTVFP = str(origiinal_MARC_data[i][35])
                VBEAZ = str(origiinal_MARC_data[i][36])
                KAUTB = str(origiinal_MARC_data[i][37])
                KORDB = str(origiinal_MARC_data[i][38])
                STAWN = str(origiinal_MARC_data[i][39])
                HERKL = str(origiinal_MARC_data[i][40])
                HERKR = str(origiinal_MARC_data[i][41])
                PRCTR = str(origiinal_MARC_data[i][42])
                SAUFT = str(origiinal_MARC_data[i][43])
                VRMOD = str(origiinal_MARC_data[i][44])
                VINT1 = str(origiinal_MARC_data[i][45])
                VINT2 = str(origiinal_MARC_data[i][46])
                VERKZ = str(origiinal_MARC_data[i][47])
                STLAL = str(origiinal_MARC_data[i][48])
                STLAN = str(origiinal_MARC_data[i][49])
                PLNNR = str(origiinal_MARC_data[i][50])
                APLAL = str(origiinal_MARC_data[i][51])
                LOSGR = str(origiinal_MARC_data[i][52])
                LGPRO = str(origiinal_MARC_data[i][53])
                DISGR = str(origiinal_MARC_data[i][54])
                RWPRO = str(origiinal_MARC_data[i][55])
                ABCIN = str(origiinal_MARC_data[i][56])
                AWSLS = str(origiinal_MARC_data[i][57])
                SERNP = str(origiinal_MARC_data[i][58])
                STDPD = str(origiinal_MARC_data[i][59])
                SFEPR = str(origiinal_MARC_data[i][60])
                RDPRF = str(origiinal_MARC_data[i][61])
                STRGR = str(origiinal_MARC_data[i][62])
                LGFSB = str(origiinal_MARC_data[i][63])
                SCHGT = str(origiinal_MARC_data[i][64])
                CCFIX = str(origiinal_MARC_data[i][65])
                EPRIO = str(origiinal_MARC_data[i][66])
                PLNTY = str(origiinal_MARC_data[i][67])
                UOMGR = str(origiinal_MARC_data[i][68])
                SFCPF = str(origiinal_MARC_data[i][69])
                SHFLG = str(origiinal_MARC_data[i][70])
                SHZET = str(origiinal_MARC_data[i][71])
                FVIDK = str(origiinal_MARC_data[i][72])
                FPRFM = str(origiinal_MARC_data[i][73])
                CASNR = str(origiinal_MARC_data[i][74])
                STEUC = str(origiinal_MARC_data[i][75])
                MATGR = str(origiinal_MARC_data[i][76])
                MINLS = str(origiinal_MARC_data[i][77])
                MAXLS = str(origiinal_MARC_data[i][78])
                FIXLS = str(origiinal_MARC_data[i][79])
                LTINC = str(origiinal_MARC_data[i][80])
                AHDIS = str(origiinal_MARC_data[i][81])
                DIBER = str(origiinal_MARC_data[i][82])
                KZPSP = str(origiinal_MARC_data[i][83])
                APOKZ = str(origiinal_MARC_data[i][84])
                LFMON = str(origiinal_MARC_data[i][85])
                LFGJA = str(origiinal_MARC_data[i][86])
                EISLO = str(origiinal_MARC_data[i][87])
                NCOST = str(origiinal_MARC_data[i][88])
                TARGET_STOCK = str(origiinal_MARC_data[i][89])
                ZZTEMPLATE = str(origiinal_MARC_data[i][90])
                ZZPRGRP = str(origiinal_MARC_data[i][91])
                LADGR = str(origiinal_MARC_data[i][92])
                ZZAPOPRGRP = str(origiinal_MARC_data[i][93])
                SOBSK = str(origiinal_MARC_data[i][94])
                QMATV = str(origiinal_MARC_data[i][95])
                WZEIT = str(origiinal_MARC_data[i][96])
                sql = (
                    r"INSERT INTO [ESS_NJ_LMDM_MD].[dbo].[TEMPMARC] ([MATNR], [WERKS], [LVORM], [MMSTA], [MMSTD], [MAABC], [EKGRP], [AUSME], [DISPR], [DISMM], [DISPO], [PLIFZ], [WEBAZ], [PERKZ], [DISLS], [BESKZ], [SOBSL], [MINBE], [EISBE], [BSTMI], [BSTMA], [BSTFE], [BSTRF], [MABST], [SBDKZ], [ALTSL], [MISKZ], [FHORI], [RGEKZ], [FEVOR], [BEARZ], [TRANZ], [DZEIT], [USEQU], [MTVFP], [VBEAZ], [KAUTB], [KORDB], [STAWN], [HERKL], [HERKR], [PRCTR], [SAUFT], [VRMOD], [VINT1], [VINT2], [VERKZ], [STLAL], [STLAN], [PLNNR], [APLAL], [LOSGR], [LGPRO], [DISGR], [RWPRO], [ABCIN], [AWSLS], [SERNP], [STDPD], [SFEPR], [RDPRF], [STRGR], [LGFSB], [SCHGT], [CCFIX], [EPRIO], [PLNTY], [UOMGR], [SFCPF], [SHFLG], [SHZET], [FVIDK], [FPRFM], [CASNR], [STEUC], [MATGR], [MINLS], [MAXLS], [FIXLS], [LTINC], [AHDIS], [DIBER], [KZPSP], [APOKZ], [LFMON], [LFGJA], [EISLO], [NCOST], [TARGET_STOCK], [ZZTEMPLATE], [ZZPRGRP], [LADGR], [ZZAPOPRGRP], [SOBSK], [QMATV], [WZEIT]) VALUES ('"
                    + MATNR
                    + r"', '"
                    + WERKS
                    + r"', '"
                    + LVORM
                    + r"', '"
                    + MMSTA
                    + r"', '"
                    + MMSTD
                    + r"', '"
                    + MAABC
                    + r"', '"
                    + EKGRP
                    + r"', '"
                    + AUSME
                    + r"', '"
                    + DISPR
                    + r"', '"
                    + DISMM
                    + r"', '"
                    + DISPO
                    + r"', '"
                    + PLIFZ
                    + r"', '"
                    + WEBAZ
                    + r"', '"
                    + PERKZ
                    + r"', '"
                    + DISLS
                    + r"', '"
                    + BESKZ
                    + r"', '"
                    + SOBSL
                    + r"', '"
                    + MINBE
                    + r"', '"
                    + EISBE
                    + r"', '"
                    + BSTMI
                    + r"', '"
                    + BSTMA
                    + r"', '"
                    + BSTFE
                    + r"', '"
                    + BSTRF
                    + r"', '"
                    + MABST
                    + r"', '"
                    + SBDKZ
                    + r"', '"
                    + ALTSL
                    + r"', '"
                    + MISKZ
                    + r"', '"
                    + FHORI
                    + r"', '"
                    + RGEKZ
                    + r"', '"
                    + FEVOR
                    + r"', '"
                    + BEARZ
                    + r"', '"
                    + TRANZ
                    + r"', '"
                    + DZEIT
                    + r"', '"
                    + USEQU
                    + r"', '"
                    + MTVFP
                    + r"', '"
                    + VBEAZ
                    + r"', '"
                    + KAUTB
                    + r"', '"
                    + KORDB
                    + r"', '"
                    + STAWN
                    + r"', '"
                    + HERKL
                    + r"', '"
                    + HERKR
                    + r"', '"
                    + PRCTR
                    + r"', '"
                    + SAUFT
                    + r"', '"
                    + VRMOD
                    + r"', '"
                    + VINT1
                    + r"', '"
                    + VINT2
                    + r"', '"
                    + VERKZ
                    + r"', '"
                    + STLAL
                    + r"', '"
                    + STLAN
                    + r"', '"
                    + PLNNR
                    + r"', '"
                    + APLAL
                    + r"', '"
                    + LOSGR
                    + r"', '"
                    + LGPRO
                    + r"', '"
                    + DISGR
                    + r"', '"
                    + RWPRO
                    + r"', '"
                    + ABCIN
                    + r"', '"
                    + AWSLS
                    + r"', '"
                    + SERNP
                    + r"', '"
                    + STDPD
                    + r"', '"
                    + SFEPR
                    + r"', '"
                    + RDPRF
                    + r"', '"
                    + STRGR
                    + r"', '"
                    + LGFSB
                    + r"', '"
                    + SCHGT
                    + r"', '"
                    + CCFIX
                    + r"', '"
                    + EPRIO
                    + r"', '"
                    + PLNTY
                    + r"', '"
                    + UOMGR
                    + r"', '"
                    + SFCPF
                    + r"', '"
                    + SHFLG
                    + r"', '"
                    + SHZET
                    + r"', '"
                    + FVIDK
                    + r"', '"
                    + FPRFM
                    + r"', '"
                    + CASNR
                    + r"', '"
                    + STEUC
                    + r"', '"
                    + MATGR
                    + r"', '"
                    + MINLS
                    + r"', '"
                    + MAXLS
                    + r"', '"
                    + FIXLS
                    + r"', '"
                    + LTINC
                    + r"', '"
                    + AHDIS
                    + r"', '"
                    + DIBER
                    + r"', '"
                    + KZPSP
                    + r"', '"
                    + APOKZ
                    + r"', '"
                    + LFMON
                    + r"', '"
                    + LFGJA
                    + r"', '"
                    + EISLO
                    + r"', '"
                    + NCOST
                    + r"', '"
                    + TARGET_STOCK
                    + r"', '"
                    + ZZTEMPLATE
                    + r"', '"
                    + ZZPRGRP
                    + r"', '"
                    + LADGR
                    + r"', '"
                    + ZZAPOPRGRP
                    + r"', '"
                    + SOBSK
                    + r"', '"
                    + QMATV
                    + r"', '"
                    + WZEIT
                    + r"')"
                )
                self.logger.info("MARC Insert SQL is" + sql)
                self.sqlList.append(sql)
        self.runSQL(self.sqlList)
        self.sqlList.clear()
        self.logger.info("Do backup and generate tables starts")
        print("Do backup and generate tables " + str(datetime.datetime.now()))
        self.runSQL([], "", "OP_MARAMARC")  # Run Procedure to do backup and generate target MARAMRAC tables

    def ExportBOM(self):
        if self.SAPConn is True:
            self.logger.info("Export BOM report starts")
            print("Export BOM report starts at " + str(datetime.datetime.now()))
            session = self.session
        else:
            self.logger.warning("SAP disconnected when exporting data when export BOM")
            print("SAP disconnected when exporting data")
            return
        self.runSQL([], "", "Insert_New_ASO")
        errorMATNR = self.runSQL([], "[ESS_NJ_LMDM_MD].[dbo].[ASO_Check_MATNR]")
        self.logger.info("Export BOM report starts")
        print("Get error material from SQL start at " + str(datetime.datetime.now()))
        errorMATNRText = r"C:\New QC\PDM\errorMATNRText.txt"
        if os.path.exists(errorMATNRText):
            os.remove(errorMATNRText)
        self.setTxtFile(errorMATNRText, errorMATNR)

        # >Start to check ASO BOM Error parts<
        self.logger.info("Export error BOM starts")
        print("Export error BOM start at " + str(datetime.datetime.now()))
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZZMRMD090"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "2491"
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"C:\New QC\PDM"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "errorMATNRText.txt"
        session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/radFL_ITEM").select()
        session.findById("wnd[0]/usr/ctxtS_STLAN-LOW").text = "5"
        session.findById("wnd[0]/usr/radFL_ITEM").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\New QC\PDM"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DailyBOMError.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        self.logger.info("Export error BOM ends")
        self.logger.info("Upload error BOM data to SQL starts")
        print("Upload error BOM data to SQL start at " + str(datetime.datetime.now()))
        self.runSQL([], "", "Delete_ASO_DUP")
        with open(os.path.join(r"C:\New QC\PDM", "DailyBOMError.txt"), "r") as errorFile:
            for errorLine in errorFile:
                errorLine = errorLine.split("\t")
                if len(errorLine) != 0 and len(errorLine) != 1:
                    if errorLine[1] == "2491":
                        MaterialNumber = str(errorLine[2]).strip()
                        BOMComponent = str(errorLine[16]).strip()
                        AlternativeBOM = str(errorLine[4]).strip()
                        DataRecordCreatedOn = str(errorLine[6]).strip()
                        SQL = (
                            r"INSERT INTO [ESS_NJ_LMDM_MD].[dbo].[ASO_Duplicate] ([Material number], [BOM component], [Alternative BOM], [Data record created on]) VALUES ('"
                            + MaterialNumber
                            + r"', '"
                            + BOMComponent
                            + r"', '"
                            + AlternativeBOM
                            + r"', '"
                            + DataRecordCreatedOn
                            + r"')"
                        )
                        self.logger.info("Upload error BOM SQL" + SQL)
                        self.sqlList.append(SQL)
            self.runSQL(self.sqlList)
            self.sqlList.clear()

        # Start to do yesterday BOM components chceck <
        self.logger.info("Export yesterday error BOM starts")
        print("Export yesterday error BOM start at " + str(datetime.datetime.now()))
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZZMRMD090"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/radFL_ITEM").select()
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "2491"
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtS_STLAN-LOW").text = "5"
        if datetime.datetime.now().weekday() == 0:
            targetDate = (datetime.datetime.today() + datetime.timedelta(-2)).strftime("%d.%m.%Y")
        else:
            targetDate = (datetime.datetime.today() + datetime.timedelta(-1)).strftime("%d.%m.%Y")
        session.findById("wnd[0]/usr/ctxtS_ANDAT-LOW").text = targetDate
        session.findById("wnd[0]/usr/radFL_ITEM").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\New QC\PDM"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "YesterBOMError.txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        self.logger.info("Export yesterday error BOM ends")
        self.logger.info("Upload yesterday error BOM data to SQL starts")
        print("Upload yesterday error BOM data to SQL start at " + str(datetime.datetime.now()))
        with open(os.path.join(r"C:\New QC\PDM", "YesterBOMError.txt"), "r") as YestErrorFile:
            for YestErrorLine in YestErrorFile:
                YestErrorLine = YestErrorLine.split("\t")
                if len(YestErrorLine) != 0 and len(YestErrorLine) != 1:
                    if YestErrorLine[1] == "2491":
                        MaterialNumber = str(YestErrorLine[2]).strip()
                        BOMComponent = str(YestErrorLine[16]).strip()
                        AlternativeBOM = str(YestErrorLine[4]).strip()
                        DataRecordCreatedOn = str(YestErrorLine[6]).strip()
                        SQL = (
                            r"DELETE FROM [ESS_NJ_LMDM_MD].[dbo].[ASO_Duplicate] WHERE [Material number] = '"
                            + MaterialNumber
                            + r"' and [BOM component] = '"
                            + BOMComponent
                            + r"';"
                            + r"INSERT INTO [ESS_NJ_LMDM_MD].[dbo].[ASO_Duplicate] ([Material number], [BOM component], [Alternative BOM], [Data record created on]) VALUES ('"
                            + MaterialNumber
                            + r"', '"
                            + BOMComponent
                            + r"', '"
                            + AlternativeBOM
                            + r"', '"
                            + DataRecordCreatedOn
                            + r"')"
                            # Add deletion firstly to avoid insert data doubly
                            # r"INSERT INTO [ESS_NJ_LMDM_MD].[dbo].[ASO_Duplicate] ([Material number], [BOM component], [Alternative BOM], [Data record created on]) VALUES ('"
                            # + MaterialNumber
                            # + r"', '"
                            # + BOMComponent
                            # + r"', '"
                            # + AlternativeBOM
                            # + r"', '"
                            # + DataRecordCreatedOn
                            # + r"')"
                        )
                        self.logger.info("Upload yesterday error BOM SQL" + SQL)
                        self.sqlList.append(SQL)
            self.runSQL(self.sqlList)
            self.sqlList.clear()
        self.runSQL([], "", "[dbo].[OP_ASO]")

    def runSQL(self, SQL=[], tableName="", procName=""):
        with pymssql.connect(
            server=self.server, database=self.database, user=self.user, password=self.password
        ) as conn:
            with conn.cursor() as cursor:
                returnResult = []
                if len(SQL) != 0:
                    for sigSQL in SQL:
                        cursor.execute(sigSQL)
                        self.logger.info("Execute SQL :" + sigSQL)
                        # print("Execute SQL:" + sigSQL)
                    conn.commit()
                    self.logger.info("SQL List has been committed")
                    print("SQL List has been done!" + str(datetime.datetime.now()))
                if str(tableName) != "":
                    sql = "SELECT * FROM " + str(tableName)
                    cursor.execute(sql)
                    result = cursor.fetchall()
                    returnResult = result
                    self.logger.info(tableName + " has been Queried")
                    print(tableName + " has been Queried!" + str(datetime.datetime.now()))
                if str(procName) != "":
                    cursor.callproc(procName)
                    conn.commit()
                    self.logger.info(procName + " has been Proceeded!")
                    print(procName + " has been Proceeded!" + str(datetime.datetime.now()))
                return returnResult

    def setTxtFile(self, filePath, list):
        with open(filePath, "w") as txtFile:
            for line in list:
                txtFile.write(str(line[0]).strip() + "\n")

    def backup(self):
        filepath = r"C:\New QC\Backup\BAK " + str(datetime.datetime.now())[:10] + r".txt"
        rstList = []  # store the result from MARAMARC
        sql = "SELECT *, CURRENT_TIMESTAMP AS VERSION FROM [ESS_NJ_LMDM_MD].[dbo].[MARAMARC]"
        with pymssql.connect(
            server=self.server, database=self.database, user=self.user, password=self.password
        ) as conn:
            with conn.cursor() as cursor:
                cursor.execute(sql)
                row = cursor.fetchall()
                with open(filepath, "w") as txtFile:
                    txtFile.write(
                        str(
                            [
                                "MATNR",
                                "MTART",
                                "MATKL",
                                "MSTAE",
                                "ZZMVGR1",
                                "WERKS",
                                "MMSTA",
                                "LADGR",
                                "HERKL",
                                "MATGR",
                                "PRCTR",
                                "LGFSB",
                                "STRGR",
                                "Version",
                            ]
                        )
                        + "\n"
                    )
                    for l in row:
                        txtFile.write(str(list(l)) + "\n")

    def OP_PDM_Script(self):
        if datetime.datetime.now().weekday() == 0:
            print("Update PDM Process Timestamp")
            self.runSQL([], "", "OP_PDM_Script")


if __name__ == "__main__":
    print("MAIN Exe start at " + str(datetime.datetime.now()))
    iniSap = SAP("P12")
    iniSap.ConnSAP()
    iniSap.ExportMARA()
    iniSap.ExportMARC()
    iniSap.uploadMARAMRAC()
    iniSap.ExportBOM()
    iniSap.backup()
    iniSap.OP_PDM_Script()
    for x in range(8)[1:]:
        if datetime.date(datetime.datetime.now().year, datetime.datetime.now().month, x).weekday() not in [5, 6]:
            standDate = datetime.date(datetime.datetime.now().year, datetime.datetime.now().month, x)
            break
    if (
        datetime.date(datetime.datetime.now().year, datetime.datetime.now().month, datetime.datetime.now().day)
        == standDate
    ):
        mail = OpOutlook.Email()
        mail.read_mail()
    print("All procedures are done")
