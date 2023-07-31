# Written for PowerShell v7.2

$SlideCharMax = 1080

$projectFolder = "C:\Script"
$SlidesFolder = "C:\Script\ARCHIVE"
$SoundFolder = "C:\Script\POLLY"
$VideosFolder = "C:\Script\VIDEOS\"

$textFolder = "$projectFolder/Text"
$introTemplate = "$projectFolder/template/Level0.Intro.v1"
$slideL0Template =  "$projectFolder/template/Level0.SelfPost.v1.pptx"
$slideL0_1Template =  "$projectFolder/template/Level0.1.SelfPost.v1.pptx"
$slideL1Template =  "$projectFolder/template/Level1v1.pptx"
$slideL1_1Template =  "$projectFolder/template/Level1.1v1.pptx"
$slideL2Template =  "$projectFolder/template/Level2v1.pptx"
$slideL2_1Template =  "$projectFolder/template/Level2.1v1.pptx"
$slideL3Template =  "$projectFolder/template/Level3v1.pptx"
$slideL3_1Template =  "$projectFolder/template/Level3.1v1.pptx"
$ExitTemplate =  "$projectFolder/template/Exitv1.pptx"
$AvatarFolder = "$projectFolder/Avatar"
$DiscOpen = "$projectFolder/template/Openv3.txt"
$DiscClose = "$projectFolder/template/Closev1.txt"
$UnixTime = get-date -UFormat %s
$SpacerSound = "$projectFolder/template/SpacerSoundv1.mp3"
$StorySpacerSound = "$projectFolder/template/StorySpacerSoundv1.mp3"
$log = "$projectFolder/log.txt"
$TrimSoundLength = 800
$global:charCounter = 0
$global:skipCharCounter = 0
$global:Voice = "Nothing"
$Disc = "$SlidesFolder/${UnixTime}.txt"
$global:VoiceHash = @{}
$global:AvatarPicture = @{}
$global:LastVoice = ""
$FinishedSlides = "$SlidesFolder/Story${UnixTime}.pptx"
[string]$audio_speed = "1.10"
$TEMP_FOLDER = "C:\AutomationMachine\TEMP"

$introTemplate = dir "$introTemplate.*"
$introTemplate = $introTemplate.FullName
$introTemplate = $introTemplate | get-random

"Starting script at $(get-date)" | out-file -append $log

function Increase-AudioSpeed
{
    param([string]$input_file)
    
    #Removed due to causing issues with truncation.
    #remove-item $folder/$hash.$audio_speed.mp3 | Out-Null
    #start-process "$PSScriptRoot\ffmpeg.exe" -ArgumentList "-i $input_file -filter:a atempo=$audio_speed $folder/$hash.$audio_speed.mp3" -NoNewWindow -Wait | Out-Null
    #"$folder/$hash.$audio_speed.mp3"
    $input_file
}


filter Remove-Curse
{
    $local:text = $_
    $local:text = $local:text.replace("fuck","#$@!%")
    $local:text = $local:text.replace("Fuck","#$@!%")
    $local:text = $local:text.replace("FUCK","#$@!%")

    $local:text

}


function StringToNumber
{
   #Turn a sentence into a number.
   param([string]$s)
   $count = 0
   foreach ($c in [char[]]$s)
   {
       $count += [int][char]$c
   }
   $count
}

function Set-Avatar 
{
    #Set $Picture varible before running
    $shapes = $slide.Shapes
    $AvatarShape =  $shapes | Where-Object {$_.name -eq "Avatar"}
    $AvatarShape.Fill.UserPicture($Picture) 2>$Null
}

function Get-StringHash
{
    param([string]$string)
    $stringAsStream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stringAsStream)
    #Avoid mixing up voices when saying same thing.
    $writer.write("$string"+$global:Voice)
    $writer.Flush()
    $stringAsStream.Position = 0
    (Get-FileHash -InputStream $stringAsStream | Select-Object Hash).hash
}

function Get-SlideTemplate
{
    #Comment level must be set
    Switch ($CommentLevel)
    {
            0     {$slideTemplate = $slideL0Template; break}
            0.1  {$slideTemplate = $slideL0_1Template; break}
            1      {$slideTemplate = $slideL1Template; break}
            1.1  {$slideTemplate = $slideL1_1Template; break}
            2      {$slideTemplate = $slideL2Template; break}
            2.1  {$slideTemplate = $slideL2_1Template; break}
            3     {$slideTemplate = $slideL3Template; break}
            3.1 {$slideTemplate = $slideL3_1Template; break}
            default {throw "Failed to switch for a slide template!"; break} 
    }
    $slideTemplate
}

function Author2Avatar
{
    param ([string]$Author)
    if ($global:AvatarPicture["$Author"])
    {
        $global:AvatarPicture["$Author"]
    }
    else
   {
      #Gather Avatars
       $AvatarFiles = dir $AvatarFolder/*
       $AvatarFiles = $AvatarFiles.FullName
       $picture = $AvatarFiles | get-random
       $global:AvatarPicture["$Author"] = $picture
       $picture
   }
}

function Author2Voice
{
    param ([string]$Author)
    if ($global:VoiceHash["$Author"])
    {
        $global:VoiceHash["$Author"]
    }
    else
   {
	 #https://docs.aws.amazon.com/polly/latest/dg/voicelist.html
       $voices = ('Amy','Emma','Brian','Aria','Ayanda',,'Joanna','Kendra','Kimberly','Salli','Joey','Matthew')
       
       #Remove Last voice to avoid duplicate voices
       $voices = foreach ($voice in $voices)
       {
           if ($voice -ne $global:LastVoice)
           {
               $Voice
           }
       }
       
       $voice = $voices | get-random
       $global:VoiceHash["$Author"] = $voice
       $global:LastVoice = $voice
       $voice
       
   }
}

function StorySpacerSound
{
        #add static spacer sound between stories
        #$slide = $presentation.Slides.Item($SlideNum )
        #$sound = $slide.Shapes.AddMediaObject2("$StorySpacerSound", 5, 5, 100, 100)
        #change to play in click sequence
        #$sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
        #$sound.Top = -100
}

function SpacerSound
{
        #insert additional spacing sound
        $slide = $presentation.Slides.Item($SlideNum )
        $sound = $slide.Shapes.AddMediaObject2("$SpacerSound", 5, 5, 100, 100)
        #change to play in click sequence
        $sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
        $sound.Top = -100
}

function Get-TextToSpeech
{
    param([string]$string, [string]$folder, [string]$log, [string]$Author)
    
   $global:Voice = Author2Voice $Author
   write-host "$Author"


    #Double quotes break Polly
    $string = $string -replace '"',"'"


    ###################################################################
    #Case Insensative
    $string = $string -replace 'fuck',"duck"
    $string = $string -replace 'bitch','witch'
    $string = $string -replace 'damn','dang'
        
    #Case Sensative
    $string = ($string).replace('TIL','Today I Learned')
    $string = ($string).replace('OP','O P')
    $string = ($string).replace('WTF', 'W T F')
    $string = ($string).replace('PTSD', 'P T S D')
    $string = ($string).replace('WWII', 'World War Two')
    $string = ($string).replace('WW2', 'World War Two')
    $string = ($string).replace('IIRC', 'if I recall correctly')
    $string = ($string).replace('AFAIK', 'as far as I know')
    $string = ($string).replace('IDGAF', "I Don't Give a Duck")
    $string = ($string).replace('GTFO', "Get the Duck Out")
    $string = ($string).replace('TL;DR', "T L D R")
    ####################################################################
 
    $hash = Get-StringHash $string
    #If we've already processed that line, reuse the artifacts
    if (Test-Path "$folder/$hash.mp3")
    {
        $global:skipCharCounter += $string.length
        start-sleep 1
        return Increase-AudioSpeed "$folder/$hash.mp3"
    }
    else
    {    
        #AWS Polly
        #aws "polly" "synthesize-speech" "--output-format" "mp3" "--engine" "neural" "--voice-id" "$global:Voice" "--text" "$string" "$folder/$hash.mp3" | Out-Null
        aws "polly" "synthesize-speech" "--region" "US-EAST-1" "--output-format" "mp3" "--voice-id" "Brian" "--text" "$string" "$folder/$hash.mp3" "--region" "us-east-2" | Out-Null
        
        
        $global:charCounter += $string.length
        
        if (Test-Path "$folder/$hash.mp3")
        {
            "Generated $hash.mp3" | out-file -append $log
            return Increase-AudioSpeed "$folder/$hash.mp3"
        }
        else
        {
            #Provide an error when the script Fails
            "The script failed to synthesize speech for the following:" | out-file -append $log
            "$string" | out-file -append $log
            start-sleep 30
            exit
        }
    }
}

function SlideShowTransition
{
    #When this is true, it's a new section.
    if ($FollowAlong -eq ($line + " "))
    {
        write-host "Applying Slide Show Transition"
        $slide.SlideShowTransition.EntryEffect = 3855
        $slide.SlideShowTransition.Duration = .5
    }
}


#Gather the Stories
$story = dir $textFolder/*
$story = $story.FullName
#Grab One Story
$story = $story | get-random


#Grab the story text
$storyText = (gc $story)[2 .. ((gc $story).count - 1)]
$OPName = (get-content $story)[0]
$OPQuestion = (get-content $story)[1]

Get-Random -SetSeed (StringToNumber (gc $story)[1])


#Close All PowerPoints
Stop-Process -name POWERPNT -Force


#Close All PowerPoints########################################################
Stop-Process -name POWERPNT -Force
Start-Sleep 1

#Open a new PowerPoint with Nothing in it
$app = New-Object -ComObject powerpoint.application
$app.Visible = $true
$presentation = $app.Presentations.Add()
Start-Sleep 1

$SlideNum = 0
#Add in the title Slide from the template
$presentation.Slides.InsertFromFile("$introTemplate", $SlideNum)

$SlideNum = 2


#copy description text
$slide = $presentation.Slides.Item(1)
$notes = $slide.NotesPage.Shapes(2)
$SlideNotes = @()
$SlideNotes += get-content $DiscOpen
$SlideNotes += $OPQuestion
$SlideNotes += $storyText[0]
$SlideNotes += get-content $DiscClose
$notes.TextFrame.TextRange.Text = $($SlideNotes | Remove-Curse) -join "`n"


#Generate Sound for Second Slide
$slide = $presentation.Slides.Item(2)

#Add story to second page notes
$notes = $slide.NotesPage.Shapes(2)
$notes.TextFrame.TextRange.Text = (gc $story) -join "`n"


#Add text for the OP username & question.
$shapes = $slide.Shapes
$CreditTextbox =  $shapes | Where-Object {$_.name -eq "Credit"}
$CreditTextBox.TextEffect.Text = $OPName | Remove-Curse
$CurrentName = $OPName
$JokeTextbox =  $shapes | Where-Object {$_.name -eq "JokeText"}
$JokeTextBox.TextEffect.Text = $OPQuestion | Remove-Curse


$SoundFile = Get-TextToSpeech "$OPQuestion" "$SoundFolder" "$log" "$CurrentName"
$sound = $slide.Shapes.AddMediaObject2("$SoundFile", 5, 5, 100, 100)
#change to play in click sequence
$sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
$sound.Top = -100


#Change the Avatar
$picture = Author2Avatar $CurrentName
Set-Avatar


SpacerSound


#change date on Joke file to make recovery easier
(Get-Item $story).LastWriteTime = (Get-Date)


#Put story to the slides
$x = 0
$FollowAlong = ""


$CommentsIntro = 0
$PriorLine = ""

foreach ($line in $storyText)
{
    $x++

    #Check if line exceeds limit of 310 for the slide:
    if (($line).length -gt 310)
    {
        Write-Host "ERROR: line $x exceeds max length"
        sleep 30
        exit
    }

    if ($x -eq 1)
    {
        #This is the URL

        #Prepare CommentLevel for possible Self Text
        $CommentLevel = 0
    }
    elseif (-not $line.trim())
    {
        SpacerSound
    }
    elseif (($line + "") -eq "")
    {
        SpacerSound
    }
    elseif (($x -eq 2) -and ($line -eq "@"))
    {
        StorySpacerSound
        SpacerSound
        $CommentLevel = 1
        $x = 1
    }
    elseif ($line -eq "@OP")
    {
        $CommentLevel = 1
        $x = 1
        StorySpacerSound
        SpacerSound
    }
    elseif ($line -eq "#OP")
    {
        #Reply from OP
        $CommentLevel = 2
        $x = 1
        SpacerSound
    }
    elseif ($line -eq "&OP")
    {
        #Reply from OP
        $CommentLevel = 3
        $x = 1
        SpacerSound
    }
    elseif ($line -eq "@")
    {
        $CommentLevel = 1
        $x = 1
        StorySpacerSound
        SpacerSound
    }
    elseif ($line -eq "#")
    {
        $CommentLevel = 2
        $x = 1
        SpacerSound
    }
    elseif ($line -eq "&")
    {
        $CommentLevel = 3
        $x = 1
        SpacerSound
    }
    elseif ($x -eq 2)
    {
        #Pick the proper Slide Template based on comment level
        $slideTemplate = Get-SlideTemplate
       
        #Setup 1st Slide of the Story
        $presentation.Slides.InsertFromFile("$slideTemplate", $SlideNum)
        $SlideNum++
       
        #Setup Credit on First Slide
        $slide = $presentation.Slides.Item($SlideNum)
        $shapes = $slide.Shapes
        $CreditTextbox =  $shapes | Where-Object {$_.name -eq "Credit"}
        $CreditTextBox.TextEffect.Text = $line | Remove-Curse
        $CurrentName = $line

        #Remove the ipsum text
        $JokeTextbox =  $shapes | Where-Object {$_.name -eq "JokeText"}
        $JokeTextBox.TextEffect.Text = " "
        
        #If the OP, use OP Name Style
        if ($OPName -eq $CreditTextBox.TextEffect.Text)
        {
            $CreditTextBox.TextEffect.Text = $line + "   OP"
            #Change Font Blue
            $CreditTextBox.TextFrame.TextRange.Font.Color.RGB = 12874308
        }

        #Hide Fuck in username
        $CreditTextBox.TextEffect.Text = $CurrentName | Remove-Curse
        
        #Set the saying preamble for use in the 3rd pass
        $saying = ""
       
       
       
        #########3Comments Intro Section############################
        if ((get-content $story)[3] -eq "@")
        {
            #If no self-text, never say comments intro
            $CommentsIntro = 2
        }
        if ($CommentsIntro -eq 1)
        {
            $saying += "Now to the comments.    "
        }
        $CommentsIntro++
        ############################################################
        
        
        if ($($PriorLine + "FILLER").substring(0,5) -eq "http:")
        {
            #Nothing
        }
        if (($PriorLine -eq "#OP") -or ($PriorLine -eq "@OP") -or ($PriorLine -eq "&OP"))
        {
            $saying += "O P replied. "
        }
        if ($PriorLine -eq "#")
        {
            #Nothing
        }
    }
    elseif ($x -eq 3)
    {
        $slide = $presentation.Slides.Item($SlideNum )
        #insert first joke line
        $shapes = $slide.Shapes
        $JokeTextbox =  $shapes | Where-Object {$_.name -eq "JokeText"}
        $JokeTextBox.TextEffect.Text = $line | Remove-Curse
        

        #Clear and Capture for the follow along effect
        $FollowAlong = $line + " "
        
        #Only works if follow along equals line
        SlideShowTransition

        #Append and Clear the saying, if used.
        $line = $saying + $line
        $saying = ""
        
        #generate sound file
        $SoundFile = Get-TextToSpeech "$line" "$SoundFolder" "$log" "$CurrentName"
        #load sound file
        $sound = $slide.Shapes.AddMediaObject2("$SoundFile", 5, 5, 100, 100)
        #change to play in click sequence
        $sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
        $sound.Top = -100
        
        #Change the Avatar
        $picture = Author2Avatar $CurrentName
        Set-Avatar
      

	    #Trim the sound length
        #$sound.MediaFormat.Endpoint = $sound.MediaFormat.Endpoint - $TrimSoundLength
    }
    elseif (($line + "") -eq "")
    {
        SpacerSound
    }
    elseif ($line -eq "$")
    {
        #Clear the follow along for a fresh slide
        $FollowAlong = ""
        #Set for dot one slides
        if (($CommentLevel % 1) -eq 0)
        {
            $CommentLevel += 0.1
        }
    }
    else
    {
        #Pick the proper Slide Template based on comment level
        $slideTemplate = Get-SlideTemplate

        #Setup the next slide
        $presentation.Slides.InsertFromFile("$slideTemplate", $SlideNum)
        $SlideNum++
        $slide = $presentation.Slides.Item($SlideNum )
        
        #insert next joke line
        $shapes = $slide.Shapes
        $JokeTextbox =  $shapes | Where-Object {$_.name -eq "JokeText"}
        $JokeTextBox.TextEffect.Text = $FollowAlong + $line | Remove-Curse
        $FollowAlong += $line + " "
        
        #Only works if follow along equals line
        SlideShowTransition
        
        #check for text overflow
        if ($FollowAlong.length -gt $SlideCharMax)
        {
            write-error "Exceeded SlideCharMax with $($FollowAlong.length) characters when $SlideCharMax is the limit."
            sleep 10
            exit
        }
        
        #Add credit
        $CreditTextbox =  $shapes | Where-Object {$_.name -eq "Credit"} 2>$Null
        $CreditTextBox.TextEffect.Text = $CurrentName
        
        #If the OP, use OP Name Style
        if ($OPName -eq $CreditTextBox.TextEffect.Text)
        {
            $CreditTextBox.TextEffect.Text = $OPName + "   OP"
            #Change Font Blue
            $CreditTextBox.TextFrame.TextRange.Font.Color.RGB = 12874308
        }

        #Hide Fuck in username
        $CreditTextBox.TextEffect.Text = $CurrentName | Remove-Curse

        #generate sound file
        $SoundFile = Get-TextToSpeech "$line" "$SoundFolder" "$log" "$CurrentName"
        #load sound file
        $sound = $slide.Shapes.AddMediaObject2("$SoundFile", 5, 5, 100, 100)
        #change to play in click sequence
        $sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
        $sound.Top = -100
         
         #Change the Avatar
         $picture = Author2Avatar $CurrentName
         Set-Avatar
         
	    #Trim the sound length
        #$sound.MediaFormat.Endpoint = $sound.MediaFormat.Endpoint - $TrimSoundLength
    }
    
    $PriorLine = $line
    
}


#add static spacer sound end of last story
$slide = $presentation.Slides.Item($SlideNum )
$sound = $slide.Shapes.AddMediaObject2("$StorySpacerSound", 5, 5, 100, 100)
#change to play in click sequence
$sound.AnimationSettings.PlaySettings.PlayOnEntry = $true
$sound.Top = -100


#add the closing slide
$presentation.Slides.InsertFromFile("$ExitTemplate", $SlideNum)


#save the slide to the project folder
$presentation.SaveAs($FinishedSlides)


#EXPORT THE VIDEO
"$SlidesFolder/Story${UnixTime}.pptx"

#Full name of slide to process
$SlidesFullPath = (get-childitem $FinishedSlides).fullname
# Name of video to process
$VideoNames = (get-childitem $FinishedSlides).name
# What to name video
$VideosFullPath = ("$VideosFolder" + "$VideoNames" + ".mp4").Replace(".pptx","")
#What to name description
$TextFullPath = ("$VideosFolder" + "$VideoNames" + ".txt").Replace(".pptx","")
#What to name thumbnail
$ThumbnailFullPath = ("$VideosFolder" + "$VideoNames" + ".png").Replace(".pptx","")
#Other info for upload
$DatFullPath = ("$VideosFolder" + "$VideoNames" + ".dat").Replace(".pptx","")
#Counting Variable for which video to process

#DEBUG
"Slides Full Path is $SlidesFullPath"
"VideoNames is $VideoNames"


$slides = $SlidesFullPath

$SKIP = $False

write-host "The current file is:  $slides"
    
#Open a new PowerPoint with Nothing in it

"Videos full path will be: $VideosFullPath"

$presentation.CreateVideo($VideosFullPath, $false, 3, 1080, 60, 100)
sleep 15

#Wait for the export to complete...
while ($(dir $VideosFullPath).length -eq 0)
{
    sleep 5
    write-host "waiting..."
    if (!(Test-Path $VideosFullPath))
    {
        sleep 5
        if (!(Test-Path $VideosFullPath[$X]))
        {
            $presentation.Close()
            Stop-Process -name POWERPNT -Force
            Start-Sleep 5
            Rename-Item -Path "$slides" -NewName "$slides.FAILED"
            $SKIP = $True
            Break
        }
    }
}

if ($SKIP)
{
    Write-Host "Exiting in 30 seconds."
    sleep 30
    Exit
}
    
#copy description text#########################################################
$slide = $presentation.Slides.Item(1)
write-host "Exporting text to $($TextFullPath)"
$($slide.NotesPage.Shapes)[1].TextEffect.Text | out-file -FilePath $TextFullPath

    
#Grab the thumbnail############################################################
$slide = $presentation.Slides.Item(2)
$shapes = $slide.Shapes
    
#Change Background to Black
$Background = $shapes | where-object {$_.name -eq "BACKGROUND"}
$OriginalBackgroundColor = $Background.Fill.ForeColor.RGB
$Background.Fill.ForeColor.RGB = 0
    
#Adjust the Font
$Title = $shapes | where-object {$_.name -eq "JokeText"}
$Length = ($Title.TextEffect.Text).length

$OriginalFontSize = $Title.TextEffect.FontSize
$Title.TextEffect.FontSize = 32
if ($Length -lt 192) {$Title.TextEffect.FontSize = 38}
if ($Length -lt 150) {$Title.TextEffect.FontSize = 44}
if ($Length -lt 120) {$Title.TextEffect.FontSize = 48}
if ($Length -lt 100) {$Title.TextEffect.FontSize = 50}
if ($Length -lt  60) {$Title.TextEffect.FontSize = 60}
   
#Save as PNG
write-host "Exporting PNG to $($ThumbnailFullPath)"
$slide.Export($ThumbnailFullPath, "PNG")
Start-Sleep 10
    
#Return Slide to Original
$Title.TextEffect.FontSize = $OriginalFontSize
$Background.Fill.ForeColor.RGB = $OriginalBackgroundColor
$presentation.Save() 

    
#Generate Upload Metadata########################################################
#Grab the Story from Slide 2
$slide = $presentation.Slides.Item(2)
$Story = $($slide.NotesPage.Shapes)[1].TextEffect.Text -split '\n'

   
#Grab the Sub Reddit Name
write-host "Grab Subreddit Name."
$URL = $Story[2]
$URL -match '.*/r/(?<Channel>.*)/comments.*'
$Channel = $Matches.Channel
   
$TitleChannel = $Channel 
  
$Channel = $Channel.Replace("MaliciousCompliance","Malicious Compliance")
$Channel = $Channel.Replace("AskReddit","Ask Reddit")
   
$Channel | out-file -FilePath $DatFullPath

#Grab the Title
write-host "Grab Title."
$Title = $Story[1]
if (! ($Title -le 100))
{
    #Truncate to 100 chars for Youtube
    $Title = $Title[0..99] -join ""
}
$Title | out-file -Append -FilePath  $DatFullPath
    
#########################################################################
$X++


#Close All PowerPoints########################################################
$presentation.Close()
Start-Sleep 20
Stop-Process -name POWERPNT -Force
Start-Sleep 1

copy-item -path "$slides" -dest "$VideosFolder"

#Clean Temp Folder#######################################################
remove-item $TEMP_FOLDER\*

#Split the Video##########################################################
write-host "Split the video into segments."
$index = 0
$num = 100
$increase = 60
while ($True){
    start-process "$PSScriptRoot\ffmpeg.exe" -ArgumentList "-i $VideosFullPath -ss $index -t $increase  $TEMP_FOLDER\$num.mp4 " -NoNewWindow -Wait | Out-Null
    if ((dir $TEMP_FOLDER\$num.mp4).length -lt 1000){
        Break
    }
    $num++
    $index += $increase
}

#Speed up the videos######################################################
write-host "Speed up video segments"

$videos = dir $TEMP_FOLDER\*
$videos = @($videos.FullName)
foreach ($video in $videos)
{
    start-process "$PSScriptRoot\ffmpeg.exe" -ArgumentList "-i $video -filter_complex `"[0:v]setpts=0.9*PTS[v];[0:a]atempo=1.10[a]`" -map `"[v]`" -map `"[a]`" $video.X.mp4" -NoNewWindow -Wait | Out-Null
}

#Combine video segments##################################################
write-host "Combine video segments"
cd $TEMP_FOLDER
# Create a list
$videos = dir $TEMP_FOLDER\*.X.mp4
$videos = @($videos.FullName)
foreach ($video in $videos)
{
    "file '$video'" | out-file -Append -FilePath  $TEMP_FOLDER\list.txt
}
# Run concat with output to video folder
$OutputVideosFullPath = $VideosFullPath.Replace(".mp4","")
$OutputVideosFullPath += "ZZZ.mp4"

start-process "$PSScriptRoot\ffmpeg.exe" -ArgumentList "-safe 0 -f concat -i list.txt -c copy $OutputVideosFullPath" -NoNewWindow -Wait | Out-Null


"$global:charCounter sent to TTS" | out-file -append $log
"$global:skipCharCounter was already on hand" | out-file -append $log
"$($global:charCounter + $global:skipCharCounter) total characters" | out-file -append $log
"Stopping script at $(get-date)" | out-file -append $log