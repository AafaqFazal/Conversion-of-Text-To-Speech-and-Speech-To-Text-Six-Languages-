
import boto3
import os
from botocore.exceptions import NoCredentialsError
import win32com.client as wincl
import speech_recognition as sr
from typing import Optional
from pydantic import BaseModel
from fastapi import FastAPI, UploadFile
import uvicorn
from datetime import datetime
from io import BytesIO
import requests
import tempfile
from textblob import TextBlob
from pydub import AudioSegment
import random

app = FastAPI()
ACCESS_KEY = 'Your Access Key'
SECRET_KEY = 'Your Secret Key'

s3 = boto3.client('s3', aws_access_key_id=ACCESS_KEY,
                  aws_secret_access_key=SECRET_KEY)
def upload_to_s3(file_name, bucket_name, s3_file_name, region):
    try:
        s3.upload_file(file_name, bucket_name, s3_file_name,
                       ExtraArgs={'ACL': 'public-read', 'ContentType': 'audio/mpeg'})
        print("Upload Successful")
        url = f"https://{bucket_name}.s3.{region}.amazonaws.com/{s3_file_name}"
        return url
    except FileNotFoundError:
        print("The file was not found")
        url = None
        return url
    except NoCredentialsError:
        print("Credentials not available")
        url = None
        return url

class SpeechToSpeechRequest(BaseModel):
    text: str
    lang: str
    gender: str
    speed: int = 1
    pitch: int = 70
    volume: int = 100
    addnoise: int = 0

@app.post("/sts")
async def Speech_to_Speech(url: str, request: SpeechToSpeechRequest):
    text = request.text
    lang = request.lang.lower()
    gender = request.gender.lower()
    speed = request.speed
    pitch = request.pitch
    volume = request.volume
    addnoise = request.addnoise
    r = sr.Recognizer()
    # with sr.WavFile(file.file) as source:
    #     audio = r.record(source)

    # Download audio file from URL and save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
        response = requests.get(url)
        f.write(response.content)
        audio_file = f.name

    # Load audio file and transcribe
    with sr.AudioFile(audio_file) as source:
        audio = r.record(source)

    if lang.lower() == "english":
        lang_code = "en-US"
    elif lang.lower() == "french":
        lang_code = "fr-FR"
    elif lang.lower() == "chinese":
        lang_code = "zh"
    elif lang.lower() == "arabic":
        lang_code = "ar-SA"
    elif lang.lower() == "italian":
        lang_code = "it-IT"
    elif lang.lower() == "japanese":
        lang_code = "ja-JP"
    elif lang.lower() == "spanish":
        lang_code = "es-ES"
    elif lang.lower() == "german":
        lang_code = "de-DE"
    elif lang.lower() == "korean":
        lang_code = "ko-KR"
    else:
        lang_code = "en-US"
    try:
        transcription = r.recognize_google(audio, language = lang_code)
        # print("Before text blob Transcription: " + transcription)
        print("Transcription: " + transcription)

        # r = str(TextBlob(transcription).correct())
        # print("after text blob Transcription: " + r)

        with open("C:\python\WindowsServerTesting-master\output.txt", "w", encoding="utf-8") as f:
            f.write(transcription) #r

        return {"transcription": transcription}
    except sr.UnknownValueError:
        print("Could not understand audio")
        return {"transcription": None}
    
    finally:
        # Delete the temporary file
        os.remove(audio_file)

    sapi = wincl.Dispatch("SAPI.SpVoice")
    sapi.Rate = speed
    sapi.Volume = volume
    time_now = str(datetime.now().timestamp())
    time_now = time_now.replace(' ', '')
    time_now = time_now.replace('.', '')
    filename = "C:/python/WindowsServerTesting-master/tts-sttAPIs/recording-{0}.wav".format(time_now)
    print(filename)
    if lang == "english":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(0)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(8)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    elif lang == "french":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(2)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(11)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    elif lang == "chinese":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(5)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(15)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    elif lang == "german":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(1)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(6)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    elif lang == "italian":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(3)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(12)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()

    elif lang == "japanese":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(4)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(13)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
            
    elif lang == "spanish":
        if gender == "male":
            sapi.Voice = sapi.GetVoices().Item(9)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
        else:
            sapi.Voice = sapi.GetVoices().Item(10)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()

    elif lang == "arabic":
            sapi.Voice = sapi.GetVoices().Item(7)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    elif lang == "korean":
            sapi.Voice = sapi.GetVoices().Item(7)
            useWincl = True
            stream = wincl.Dispatch("SAPI.SpFileStream")
            stream.Open(os.path.abspath(filename), 3, False)
            sapi.AudioOutputStream = stream
            sapi.Speak(text)
            stream.Close()
    else:
        return {"message": "Language not supported!"}
    
    if addnoise == 1:
        narration = AudioSegment.from_wav(filename)
        ambient_noise = AudioSegment.from_mp3("C:/python/WindowsServerTesting-master/noise1.mp3")

        start_time =0
        #random.randint(0, len(ambient_noise) - len(narration))
        noise_section = ambient_noise[start_time:start_time + len(narration)]

        mixed_audio = narration.overlay(noise_section)

        mixed_audio.export(filename, format="wav")

    s3_file = "recording-{0}.wav".format(time_now)
    url = upload_to_s3(filename, "covermatic-voice", s3_file, "us-east-2")
    try:
        os.remove(filename)
    except:
        pass
    return {"message": "Text converted to speech successfully!", "url": url}
if __name__ == "__main__":
    uvicorn.run("tts_stt_test:app", host="0.0.0.0", port=8000)