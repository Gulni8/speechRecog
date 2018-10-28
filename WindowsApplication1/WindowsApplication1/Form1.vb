
Imports System.Speech.Recognition
Imports AIMLbot



Public Class Form1

    Dim sp As New SpeechRecognitionEngine

    Dim mybot As New Bot
    Dim myuser As New User("CurrentUser", mybot)

    Public Sub InitRecog()
        Try
            Dim gb As New GrammarBuilder
            gb.AppendDictation()

            Dim gram As New Grammar(gb)

            sp.LoadGrammar(gram)
            sp.SetInputToDefaultAudioDevice()

            AddHandler sp.SpeechDetected, AddressOf SpeechDet
            AddHandler sp.SpeechHypothesized, AddressOf SpeechHyp
            AddHandler sp.SpeechRecognized, AddressOf SpeechRec
            AddHandler sp.SpeechRecognitionRejected, AddressOf SpeechRej


        Catch ex As Exception

            MsgBox("Error : Problem initializing Speech Recognizer", MsgBoxStyle.Critical, "-Speech Recognition System")


        End Try



    End Sub

    Public Sub InitAIML()

        Try

            mybot.loadSettings()
            mybot.isAcceptingUserInput = False
            mybot.loadAIMLFromFiles()
            mybot.isAcceptingUserInput = True


        Catch ex As Exception

            MsgBox("Error : Can't Load AIML Files", MsgBoxStyle.Critical, "AIML Files - ")

        End Try

    End Sub

    Public Function GetReply()

        Dim req As New Request(vbQuestion, myuser, mybot)
        Dim res As New Result(myuser, mybot, req)
        res = mybot.Chat(req)
        Return res.Output

    End Function



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        InitRecog()
        InitAIML()


    End Sub



    'Speech Events



    Private Sub SpeechDet(ByVal sender As Object, ByVal e As SpeechDetectedEventArgs)

        Label1.Text = "Listening. . ."

    End Sub

    Private Sub SpeechHyp(ByVal sender As Object, ByVal e As SpeechHypothesizedEventArgs)

        Button1.Text = e.Result.Text

    End Sub

    Private Sub SpeechRec(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)

        Label1.Text = "I heard you"

        Dim response As String

        response = GetReply(e.Result.Text)

        TextBox1.AppendText("You > " & e.Result.Text & vbCrLf)
        TextBox1.AppendText("-- > " & response & vbCrLf)


        StopRecog()
        Button1.Text = "Speak"

        Button1.Enabled = True


    End Sub

    Private Sub SpeechRej(ByVal sender As Object, ByVal e As SpeechRecognitionRejectedEventArgs)

        Label1.Text = "can't understand what you said"


    End Sub


    Public Sub StopRecog()

        Try
            sp.RecognizeAsyncStop()

        Catch ex As Exception

        End Try
    End Sub


    


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Button1.Enabled = False

        sp.RecognizeAsync(RecognizeMode.Multiple)

    End Sub

End Class
