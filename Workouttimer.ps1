## Variables ##
$endtime = (get-date).AddMinutes(10)
$startexercise = "Start"
$breakexercise = "Rest"
$finished = "All Done"

## Add Speech stuff ##
Add-Type -AssemblyName System.Speech
$Speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
$Speech.SelectVoice("Microsoft Zira Desktop")

## Meat and Potatoes ##
While ($endtime -gt (get-date)){
    $Speech.Speak($startexercise)
    Sleep -Seconds 20
    $Speech.Speak($breakexercise)
    Sleep -Seconds 10
    }
$Speech.Speak($finished)

