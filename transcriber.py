import sys  # core python module
from pathlib import Path  # core python module
from time import sleep  # core python module

import requests  # pip install requests
import xlwings as xw  # pip install xlwings
from pytube import YouTube  # pip install pytube
from wordcloud import WordCloud  # pip install wordcloud


def generate_wordcloud(textfile, output_path, status_cell):
    textfile = Path(textfile)
    content = textfile.read_text()
    wordcloud = WordCloud().generate(content)
    wordcloud.to_file(Path(output_path) / f"{textfile.stem}.png")
    status_cell.value = "WordCloud generated"


def download_youtube_video(youtube_url, status_cell, output_path):
    # download the file
    audio_file = YouTube(youtube_url).streams.get_audio_only().download(output_path=output_path)

    # by default, the file will be saved as mp4
    # I will replace the suffix to save it as mp3
    audio_file = Path(audio_file)
    audio_file = audio_file.replace(audio_file.with_suffix(".mp3"))

    status_cell.value = f"YT-Video saved as mp3: {audio_file}"
    return audio_file


def read_file(filename, chunk_size=5242880):
    with open(filename, "rb") as _file:
        while True:
            data = _file.read(chunk_size)
            if not data:
                break
            yield data


def transcribe_audio_file(api_key, status_cell, audio_file, output_path):
    # Config headers
    headers = {"authorization": api_key, "content-type": "application/json"}

    # upload audio file to assemblyai via post request
    # For reference, an upload response will look like this:
    # {'upload_url': 'https://cdn.assemblyai.com/upload/63928cd3-152e-4024-8e28-fd7174ec0b4d'}
    status_cell.value = "Uploading audio file to AssemblyAI..."
    upload_endpoint = "https://api.assemblyai.com/v2/upload"
    upload_response = requests.post(upload_endpoint, headers=headers, data=read_file(audio_file))
    status_cell.value = "Audio file uploaded"

    # request transcription of uploaded file
    transcript_endpoint = "https://api.assemblyai.com/v2/transcript"
    transcript_request = {"audio_url": upload_response.json()["upload_url"], "language_code": "en"}
    transcript_response = requests.post(transcript_endpoint, json=transcript_request, headers=headers)
    status_cell.value = "Transcription Requested"

    # transcribing a file might take some time
    # we will keep track of the current status using our ID
    polling_response = requests.get(transcript_endpoint + "/" + transcript_response.json()["id"], headers=headers)
    while polling_response.json()["status"] != "completed":
        sleep(10)
        polling_response = requests.get(transcript_endpoint + "/" + transcript_response.json()["id"], headers=headers)
        status_cell.value = f"File is {polling_response.json()['status']}"

    # Once the transcription is completed, we will save it to an text file
    data = polling_response.json()["text"]
    transcription_txt = output_path / f"{Path(audio_file).stem}_transcription.txt"
    transcription_txt.write_text(data)
    status_cell.value = f"Transcript saved to {transcription_txt}"
    return transcription_txt


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    youtube_url = sheet["YOUTUBE_URL"].value
    api_key = sheet["API_KEY"].value
    transcribe = sheet["TRANSCRIBE"].value
    wordcloud = sheet["WORDCLOUD"].value
    status_cell = sheet["STATUS_CELL"]

    # Reset status
    status_cell.value = ""

    output_path = Path(__file__).parent

    if youtube_url:
        status_cell.value = "Downloading audio file ..."
        audio_file = download_youtube_video(youtube_url, status_cell, output_path)
    else:
        status_cell.value = "No YouTube Link entered"
        sys.exit()

    if transcribe:
        transcription_text = transcribe_audio_file(api_key, status_cell, audio_file, output_path)

    if transcribe and wordcloud:
        generate_wordcloud(transcription_text, output_path, status_cell)


if __name__ == "__main__":
    xw.Book("transcriber.xlsm").set_mock_caller()
    main()
