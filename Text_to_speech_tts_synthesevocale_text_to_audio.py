import os
import re
from google.cloud import texttospeech
from google.api_core.exceptions import InvalidArgument
from pydub import AudioSegment
import docx
from io import BytesIO
import simpleaudio as sa

# Set up the environment variable for the service account key file
api_config_path = os.path.join(os.getcwd(),"text-to-speech-440500-38ed493f3165.json")
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = api_config_path

def get_voices_supporting_speaking_rate(speaking_rate=0.5):
    client = texttospeech.TextToSpeechClient()
    voices = client.list_voices()
    french_voices = [voice for voice in voices.voices if "fr-FR" in voice.language_codes]
    voices_accepting_speaking_rate = []
    voices_not_accepting_speaking_rate = []

    for voice in french_voices:
        voice_params = texttospeech.VoiceSelectionParams(
            language_code="fr-FR",
            name=voice.name,
            ssml_gender=voice.ssml_gender
        )
        audio_config = texttospeech.AudioConfig(
            audio_encoding=texttospeech.AudioEncoding.MP3,
            speaking_rate=speaking_rate
        )

        try:
            synthesis_input = texttospeech.SynthesisInput(text="Test text for compatibility check.")
            client.synthesize_speech(input=synthesis_input, voice=voice_params, audio_config=audio_config)
            voices_accepting_speaking_rate.append(voice)
        except InvalidArgument:
            voices_not_accepting_speaking_rate.append(voice)

    return voices_accepting_speaking_rate, voices_not_accepting_speaking_rate

def create_voice_name_slowable_dict(slowable_voices, regular_voices):
    """
    Creates a dictionary to check if a voice is slowable and the appropriate gender.
    """
    voice_name_slowable_dict = {}
    for voice in slowable_voices:
        voice_name_slowable_dict[voice.name] = {"slowable": True, "gender": voice.ssml_gender}
    for voice in regular_voices:
        voice_name_slowable_dict[voice.name] = {"slowable": False, "gender": voice.ssml_gender}
    return voice_name_slowable_dict

def preprocess_text_from_file(file_path):
    """
    Reads text from a Word (.docx) file, cleans it, and returns a list of paragraphs.
    """
    paragraphs = []
    if file_path.endswith('.docx'):
        doc = docx.Document(file_path)
        
        # Extract and clean paragraphs
        paragraphs = [re.sub(r'\d+\)', '', para.text).strip() for para in doc.paragraphs if para.text.strip()]
        
        # Remove specific characters and replace symbols as needed
        paragraphs = [para.replace("[", "").replace("]", "") for para in paragraphs]
        paragraphs = [re.sub(r'[*#@]', '', para) for para in paragraphs]
        paragraphs = [para.replace("/", ", ") for para in paragraphs]
        
    else:
        raise ValueError("Unsupported file type. Please provide a .docx file.")
    
    return paragraphs

def synthesize_speech(paragraphs, output_dir, voice_name_slowable_dict, affordable_voices, prefix="French_", paragraph_pause_duration="3s", punctuation_pause_duration="1s", speaking_rate=0.85, rotate_voice = 10,play_audio=False):
    """
    Synthesizes speech for each paragraph/sentence individually, saving each as a separate audio file or playing it directly.
    Rotates voices every 'rotate_voice' sentences from affordable_voices list.
    """
    client = texttospeech.TextToSpeechClient()

    # Ensure output directory exists if saving files
    if not play_audio:
        os.makedirs(output_dir, exist_ok=True)

    # Rotate through affordable voices every 'rotate_voice' sentences
    affordable_voice_names = list(affordable_voices.keys())
    num_voices = len(affordable_voice_names)
    
    for idx, paragraph in enumerate(paragraphs, start=1):
        # Determine the current voice based on the rotate_voice-sentence rotation
        voice_idx = (idx - 1) // rotate_voice % num_voices
        current_voice_name = affordable_voice_names[voice_idx]
        current_gender = affordable_voices[current_voice_name]
        
        # Set up the voice parameters
        voice_params = texttospeech.VoiceSelectionParams(
            language_code="fr-FR",
            name=current_voice_name,
            ssml_gender=texttospeech.SsmlVoiceGender.MALE if current_gender == "Male" else texttospeech.SsmlVoiceGender.FEMALE
        )

        # Check if the voice is slowable
        use_slow_voice = voice_name_slowable_dict[current_voice_name]["slowable"]
        
        audio_config = texttospeech.AudioConfig(
            audio_encoding=texttospeech.AudioEncoding.MP3,
            speaking_rate=speaking_rate if use_slow_voice else 1.0
        )

        # Split text by punctuation for pauses
        segments = re.split(r'([.;:!?…])', paragraph)
        combined_audio = AudioSegment.empty()
        
        for i in range(0, len(segments), 2):
            segment = segments[i].strip()
            punctuation = segments[i + 1] if i + 1 < len(segments) else ""
            
            # Synthesize each segment individually
            if segment:
                try:
                    synthesis_input = texttospeech.SynthesisInput(text=segment + punctuation)
                    response = client.synthesize_speech(input=synthesis_input, voice=voice_params, audio_config=audio_config)
                    segment_audio = AudioSegment.from_file(BytesIO(response.audio_content), format="mp3")
                    combined_audio += segment_audio
                except InvalidArgument:
                    print(f"Voice '{current_voice_name}' does not support SSML or encountered an issue.")
                    continue

            # Add a pause after punctuation
            if punctuation:
                pause_duration_ms = int(float(punctuation_pause_duration.strip("s")) * 1000)
                silent_segment = AudioSegment.silent(duration=pause_duration_ms)
                combined_audio += silent_segment

        # Add a longer pause at the end of each paragraph/sentence
        paragraph_pause_ms = int(float(paragraph_pause_duration.strip("s")) * 1000)
        combined_audio += AudioSegment.silent(duration=paragraph_pause_ms)

        # Play audio directly if `play_audio` is True, else save it
        if play_audio:
            play_obj = sa.play_buffer(
                combined_audio.raw_data,
                num_channels=combined_audio.channels,
                bytes_per_sample=combined_audio.sample_width,
                sample_rate=combined_audio.frame_rate
            )
            play_obj.wait_done()
        else:
            # Save the audio file for each paragraph/sentence
            output_audio_path = os.path.join(output_dir, f"{prefix}{idx}.mp3")
            combined_audio.export(output_audio_path, format="mp3")
            print(f"Saved paragraph {idx} as '{output_audio_path}'.")

def synthesize_text_or_file(test_text=None, file_path=None, output_dir="AudioFromText", prefix="French_", paragraph_pause_duration="3s", punctuation_pause_duration="1s", speaking_rate=0.85, voice_name_slowable_dict=None, affordable_voices=None, rotate_voice = 10, play_audio=False):
    """
    Decides whether to process a single sentence (test_text) or an entire document (file_path),
    rotating through affordable voices every 'rotate_voice' sentences.
    """
    if test_text:
        paragraphs = [test_text]  # Treat single sentence as a list with one item
    elif file_path:
        paragraphs = preprocess_text_from_file(file_path)  # Process the document
    else:
        raise ValueError("Either 'test_text' or 'file_path' must be provided.")
    
    # Call the synthesis function with the option to play or save audio
    synthesize_speech(paragraphs, output_dir, voice_name_slowable_dict, affordable_voices, prefix, paragraph_pause_duration, punctuation_pause_duration, speaking_rate, rotate_voice,play_audio)

# Generate lists of slow-compatible voices and regular voices
slowable_voices, regular_voices = get_voices_supporting_speaking_rate(speaking_rate=0.5)
voice_name_slowable_dict = create_voice_name_slowable_dict(slowable_voices, regular_voices)

# Define preferred affordable voices for rotation
preferred_affordable_voices = {
    "fr-FR-Polyglot-1": "Male",
    "fr-FR-Wavenet-D": "Male",
    "fr-FR-Wavenet-C": "Female"
}

# Example usage with single test text
test_text = "Bonjour! Ceci est un test de synthèse vocale en français."
synthesize_text_or_file(test_text=test_text, output_dir="AudioFromText", prefix="French_", paragraph_pause_duration="3s", punctuation_pause_duration="1s", speaking_rate=0.85, voice_name_slowable_dict=voice_name_slowable_dict, affordable_voices=preferred_affordable_voices, play_audio=True)

# Example usage with a full document and voice rotation every 'rotate_voice' sentences
file_path = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Nufi\lecture_du_livre_phrasebook_Nufi_version_Francaise_Text_to_Speech_TTS_test.docx"

# file_path = r"C:\Users\tcham\OneDrive\Documents\testlecture.docx"

synthesize_text_or_file(file_path=file_path, output_dir="AudioFromText", prefix="French_", paragraph_pause_duration="3s", punctuation_pause_duration="1s", speaking_rate=0.85, voice_name_slowable_dict=voice_name_slowable_dict, affordable_voices=preferred_affordable_voices, rotate_voice = 10, play_audio=False)
