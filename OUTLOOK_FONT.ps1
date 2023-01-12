$outlook = New-Object -ComObject Outlook.Application
$options = $outlook.Options
$options.EmailSignature.NewMessageFont.Name = "Arial"
$options.EmailSignature.ReplyMessageFont.Name = "Arial"
$options.EmailSignature.NewMessageFont.Size = 10
$options.EmailSignature.ReplyMessageFont.Size = 10
$options.EmailSignature.NewMessageFont.Bold = $false
$options.EmailSignature.ReplyMessageFont.Bold = $false
$options.EmailSignature.NewMessageFont.Italic = $false
$options.EmailSignature.ReplyMessageFont.Italic = $false
$options.Save()