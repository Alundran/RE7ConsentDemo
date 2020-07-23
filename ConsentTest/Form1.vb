Imports Blackbaud.PIA.RE7.BBREAPI

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oApi = New REAPI
        'Enter your Serial Number, Username/Password & Database Number for the connection to RE7
        oApi.Init("WRE11111", "Supervisor", "ADMIN", 50, "", AppMode.amStandalone)
        oApi.SignOutOnTerminate = True
        TextBox1.AppendText("API initiated")
        TextBox1.AppendText(Environment.NewLine)
        Dim constit As CRecord
        TextBox1.AppendText("CRecord Initiated")
        TextBox1.AppendText(Environment.NewLine)

        'Consents are Objects
        Dim Consent As Object
        Dim constit2 As IBBRecord2

        Dim consents As Object
        constit = New CRecord
        constit.Init(oApi.SessionContext) 'Init a constituent
        constit.Load(635) 'Load Mark Adamson's System Record ID
        constit2 = constit
        TextBox1.AppendText("Loaded Mark Adamson")
        TextBox1.AppendText(Environment.NewLine)

        'Load the constituent consent collection, all the consents currently on the constituent record
        TextBox1.AppendText("Attempting ContactConsents")
        TextBox1.AppendText(Environment.NewLine)
        consents = constit2.ContactConsents


        'add a new consent record
        Consent = consents.Add
        Consent.Fields(EContactConsent_Fields.ContactConsent_fld_Category) = "Newsletter"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Channel) = "Phone"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Consent_Date) = "11/20/2020"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Consent_Statement) = "You are agreeing to be contacted"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Constit_ID) = 635
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Date_Added) = Now()
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Form_Location) = "C:\Local\My Folder"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Form_Version) = "1.2.1"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Privacy_Policy) = "This is Super Secret"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Response) = "Opt-in"
        Consent.Fields(EContactConsent_Fields.ContactConsent_fld_Source) = "API"
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Sequence) = 1
        Consent.Fields(EContactConsent_Fields.ContactConsent_Fld_Username) = "Derek Marr"

        'Validate that we are saving a valid consent
        Consent.Validate()
        'Save the consent or the entire collection
        Consent.Save()
        consents.Save()

        'Close the record
        constit.CloseDown()
        TextBox1.AppendText("Consent record added successfully")
        TextBox1.AppendText(Environment.NewLine)

    End Sub
End Class
