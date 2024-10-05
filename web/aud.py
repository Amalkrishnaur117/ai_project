from gtts import gTTS
def create_audio(script, audio_filename="presentation_audio.mp3"):
    # Create audio from script
    audio = gTTS(text=script, lang='en')
    audio.save(audio_filename)
    print(f"Audio saved as '{audio_filename}'.")

# Main execution
if __name__ == "__main__":
    # Create the PowerPoint presentation

    # Define your script for the audio
    script = """
    Welcome to this presentation on Key Points.
    First, we will discuss the introduction to the topic.
    Next, we will highlight the importance of the topic.
    Then, we will cover key findings and future implications.
    Finally, we will conclude our presentation. Thank you for your attention!
    """

    # Create the audio file from the script
    create_audio(script)