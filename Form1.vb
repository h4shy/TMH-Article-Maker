Imports System.Net
Imports OpenAI_API
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Form1

    Dim file_WH As String = Path.Combine(Application.StartupPath, "webhookURL.txt")
    Dim file_TestWH As String = Path.Combine(Application.StartupPath, "testWebhookURL.txt")
    Dim file_KeyOpenAI As String = Path.Combine(Application.StartupPath, "openaiKeyAPI.txt")
    Dim webhookURL As String = File.ReadAllText(file_WH)
    Dim testWebhookURL As String = File.ReadAllText(file_TestWH)
    Dim openaiKeyAPI As String = File.ReadAllText(file_KeyOpenAI)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim pattern_WH As String = "https:\/\/discord\.com\/api\/webhooks\/\d+\/[\w-]+"
        Dim pattern_OpenAI As String = "sk-[A-Za-z0-9]{32}"

        While Not Regex.IsMatch(webhookURL, pattern_WH)
            File.WriteAllText(file_WH, InputBox("Input your main Discord webhook URL." & vbCrLf & "This prompt will repeat if there is an error"))
            webhookURL = File.ReadAllText(file_WH)
        End While
        While Not Regex.IsMatch(testWebhookURL, pattern_WH)
            File.WriteAllText(file_TestWH, InputBox("Input your testing-only Discord webhook URL." & vbCrLf & "This prompt will repeat if there is an error"))
            testWebhookURL = File.ReadAllText(file_TestWH)
        End While
        While Not Regex.IsMatch(openaiKeyAPI, pattern_OpenAI)
            File.WriteAllText(file_KeyOpenAI, InputBox("Input your OpenAI API Key." & vbCrLf & "This prompt will repeat if there is an error"))
            openaiKeyAPI = File.ReadAllText(file_KeyOpenAI)
        End While
    End Sub

    Async Function Generate_AI(prompt As String) As Task(Of String)
        Dim api As New OpenAIAPI(openaiKeyAPI)
        Dim chatRequest As Chat.ChatRequest = New Chat.ChatRequest() With {
        .Model = Models.Model.ChatGPTTurbo,
        .Temperature = 0.5,
        .MaxTokens = 500,
        .Messages = New Chat.ChatMessage() {New Chat.ChatMessage(Chat.ChatMessageRole.User, prompt)}
        }
        ProgressBar1.Value = 1
        Dim chatResult As Chat.ChatResult = Await api.Chat.CreateChatCompletionAsync(chatRequest)
        ProgressBar1.Value = 3
        Dim comment As String = chatResult.Choices(0).Message.Content
        Return comment
    End Function

    Async Function Post_WH(message As String) As Task(Of String)
        Dim payload As String = $"{{""content"": ""{message}""}}"
        Dim client As New WebClient()
        client.Headers.Add("Content-Type", "application/json")

        If ComboBox1.Text = "main-webhook" Then
            Dim wHookURL As String = webhookURL
            Await client.UploadStringTaskAsync(wHookURL, "POST", payload)
        Else
            Dim wHookURL As String = testWebhookURL
            Await client.UploadStringTaskAsync(wHookURL, "POST", payload)
        End If
        Return payload
    End Function

    Private Async Sub GENERATE_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrEmpty(RichTextBox1.Text) Or String.IsNullOrEmpty(TextBox1.Text) Then
            MsgBox("You have to enter a message first and set an author to generate an article")
        Else
            ProgressBar1.Visible = True
            ProgressBar1.Value = 0
            Dim message As String = RichTextBox1.Text
            Dim author As String = TextBox1.Text
            Dim prompt As String =
                $"(Organize the following message and prepare it to be suitable for sharing as a brief data article on a Discord server. infrequently use Discord-friendly formatting and emojis.) Message: {message}. - {author}"

            RichTextBox1.Text = Await Generate_AI(prompt)
            Threading.Thread.Sleep(1000)
            ProgressBar1.Value = 5
            Threading.Thread.Sleep(1000)
            ProgressBar1.Value = 7
            Threading.Thread.Sleep(1000)
            ProgressBar1.Value = 10
            ProgressBar1.Visible = False
        End If
    End Sub

    Private Async Sub SEND_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If String.IsNullOrEmpty(RichTextBox1.Text) Then
            MsgBox("You have to enter a message or generate an article first")
        ElseIf RichTextBox1.Text.Length > RichTextBox1.MaxLength Then
            MsgBox("The message is way too long." & vbCrLf & $"{RichTextBox1.Text.Length} / 2000")
        Else
            Dim message As String = RichTextBox1.Text
            MessageBox.Show($"{message}", "Click OK to send the article.", MessageBoxButtons.OK)
            Dim formattedText As String = ""
            For Each c As Char In message
                If c = vbCr Or c = vbLf Then
                    formattedText &= "\n"
                Else
                    formattedText &= c
                End If
            Next
            Await Post_WH(formattedText)
            MsgBox("Done.")
            RichTextBox1.Text = ""
        End If
    End Sub

    Private Async Sub TEST_Button(sender As Object, e As EventArgs) Handles Button3.Click
        Dim prompt As String = "This is a test. Please respond to this prompt with no more than 5 words stating that the test was successful."
        If CheckBox1.Checked = False And CheckBox2.Checked = True Then
            MsgBox(Await Generate_AI(prompt))
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = True Then
            Dim comment As String = Await Generate_AI(prompt)
            Await Post_WH(comment)
            MsgBox("Done. Check if the webhook received the post." & vbCrLf & $"[{comment}]")
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = False Then
            Dim message As String = "TEST - if you see this message = the test was successful"
            MsgBox("Done. Check if the webhook received the post." & vbCrLf & $"[{Await Post_WH(message)}]")
        Else
            MsgBox("ok listen, I didn't wanna say anything but... you see the checkboxes above this button? use them.")
        End If
    End Sub

    Private Sub CopyText(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Clipboard.SetText(RichTextBox1.Text)
    End Sub

    Private Sub PasteText(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If Clipboard.ContainsText() Then
            RichTextBox1.Text = Clipboard.GetText()
        End If
    End Sub

    Private Sub Modify_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim pattern_WH As String = "https:\/\/discord\.com\/api\/webhooks\/\d+\/[\w-]+"
        Dim pattern_OpenAI As String = "sk-[A-Za-z0-9]{32}"

        Dim settingNum As String = InputBox("Which of the following you would like to modify?" & vbCrLf & "Input a number: " & vbCrLf & "1. main-webhook | 2. testing-webhook | 3. openaiKeyAPI")
        If settingNum = "1" Then
            File.WriteAllText(file_WH, InputBox("Input your main Discord webhook URL"))
            webhookURL = File.ReadAllText(file_WH)
            While Not Regex.IsMatch(webhookURL, pattern_WH)
                File.WriteAllText(file_WH, InputBox("Input your main Discord webhook URL." & vbCrLf & "This prompt will repeat if there is an error"))
                webhookURL = File.ReadAllText(file_WH)
            End While
        ElseIf settingNum = "2" Then
            File.WriteAllText(file_TestWH, InputBox("Input your testing-only Discord webhook URL"))
            testWebhookURL = File.ReadAllText(file_TestWH)
            While Not Regex.IsMatch(testWebhookURL, pattern_WH)
                File.WriteAllText(file_TestWH, InputBox("Input your testing-only Discord webhook URL." & vbCrLf & "This prompt will repeat if there is an error"))
                testWebhookURL = File.ReadAllText(file_TestWH)
            End While
        ElseIf settingNum = "3" Then
            File.WriteAllText(file_KeyOpenAI, InputBox("Input your OpenAI API Key"))
            openaiKeyAPI = File.ReadAllText(file_KeyOpenAI)
            While Not Regex.IsMatch(openaiKeyAPI, pattern_OpenAI)
                File.WriteAllText(file_KeyOpenAI, InputBox("Input your OpenAI API Key." & vbCrLf & "This prompt will repeat if there is an error"))
                openaiKeyAPI = File.ReadAllText(file_KeyOpenAI)
            End While
        End If
        settingNum = ""
    End Sub

End Class


