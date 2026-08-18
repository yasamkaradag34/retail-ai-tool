import tempfile
import os

model = None
try:
    import whisper
    model = whisper.load_model("base")
except Exception as e:
    print(f"⚠️ Whisper yüklenemedi (cloud modunda normal): {e}", flush=True)

def transcribe_audio(audio_bytes: bytes) -> str:
    if model is None:
        return "Sesli komut desteği bu ortamda aktif değil."
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
        f.write(audio_bytes)
        tmp_path = f.name

    try:
        result = model.transcribe(tmp_path, language="tr")
        return result["text"].strip()
    finally:
        os.unlink(tmp_path)