import math
from pathlib import Path
import os

import moviepy.editor as mp
from PIL import Image
import numpy as np
from openai import OpenAI
import xlwings as xw
from dotenv import load_dotenv
load_dotenv()


def read_phrases(wb_name, rng, sht_name="Sheet1"):
    wb = xw.Book(wb_name)
    sht = wb.sheets(sht_name)
    phrases = sht.range(rng)
    for p in phrases.rows:
        yield (p[1].value, p[2].value)


def text_to_mp3(text, voice, out_mp3):
    client = OpenAI(api_key=os.getenv("OPEN_AI_KEY"), base_url=os.getenv("OPEN_AI_BASE_URL"))
    response = client.audio.speech.create(model="tts-1", voice=voice, input=text)

    response.write_to_file(out_mp3)


def create_video_from_mp3(mp3_file, output_file, image_file=None, duration=None):
    # 加载音频文件
    audio = mp.AudioFileClip(mp3_file)

    # 如果没有指定持续时间，使用音频文件的长度
    if duration is None:
        duration = audio.duration

    # 创建一个纯色背景图像（如果没有提供图像文件）
    if image_file is None:
        img = Image.new("RGB", (640, 480), color=(73, 109, 137))
        img_array = np.array(img)
    else:
        img = Image.open(image_file)
        img_array = np.array(img)

    # 创建视频剪辑
    video = mp.ImageClip(img_array).set_duration(duration)

    # 将音频添加到视频
    video = video.set_audio(audio)

    # 写入输出文件
    video.write_videofile(output_file, fps=24)


if __name__ == "__main__":
    file_root = "./src/auto_tts/files"
    for i, d in enumerate(read_phrases("./src/auto_tts/phrases.xlsx", "A2:C5")):
        mp3_name = os.path.join(file_root, f"p-{i+1}.mp3")
        mp4_name = os.path.join(file_root, f"p-{i+1}.mp4")

        text_to_mp3(d[0], d[1], mp3_name)
        create_video_from_mp3(mp3_name, mp4_name)


