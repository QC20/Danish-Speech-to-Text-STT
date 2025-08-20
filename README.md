# Danish Speech-to-Text (STT): Interview Transcription Tool

A powerful Python script for transcribing research interviews using OpenAI's Whisper Large-v2 model. This tool provides high-quality, academic-grade transcriptions with enhanced speaker detection, detailed statistics, and professional document formatting.

## Features

- **High-Quality Transcription**: Uses OpenAI Whisper Large-v2 with enhanced quality parameters
- **Multi-Language Support**: Supports 50+ languages with explicit language specification
- **Smart Speaker Detection**: Intelligent pause-based speaker change detection
- **Academic Formatting**: Generates professional Word documents and text files
- **Detailed Statistics**: Comprehensive analysis of speech patterns, confidence scores, and timing
- **Context-Aware**: Uses interview context to improve transcription accuracy
- **Quality Filtering**: Automatically filters low-confidence segments

## Prerequisites

- Python 3.8 or higher
- Basic familiarity with Python and command line
- Audio file in supported format (mp3, wav, m4a, flac, etc.)

## Installation

### macOS (Tested)

1. **Install dependencies using pip:**
   ```bash
   pip install openai-whisper python-docx
   ```

2. **Alternative: Create a virtual environment (recommended):**
   ```bash
   # Create virtual environment
   python3 -m venv whisper-transcription
   
   # Activate virtual environment
   source whisper-transcription/bin/activate
   
   # Install dependencies
   pip install openai-whisper python-docx
   ```

### Windows (Untested - theoretical instructions)

1. **Install dependencies using pip:**
   ```cmd
   pip install openai-whisper python-docx
   ```

2. **Alternative: Create a virtual environment (recommended):**
   ```cmd
   # Create virtual environment
   python -m venv whisper-transcription
   
   # Activate virtual environment
   whisper-transcription\Scripts\activate
   
   # Install dependencies
   pip install openai-whisper python-docx
   ```

### First-Time Setup

The first time you run the script, Whisper will download the Large-v2 model (~3GB). This happens automatically but requires an internet connection and may take several minutes depending on your connection speed.

## Quick Start

1. **Download the script** and save it as `interview_transcription.py`

2. **Open the script** in your preferred text editor

3. **Configure the settings** in the `main()` function (around line 300):

   ```python
   # AUDIO FILE SETTINGS
   audio_file_path = "/path/to/your/audio/file.mp3"
   
   # PARTICIPANT INFORMATION
   interviewer_name = "Interviewer (Dr. Smith)"
   participant_name = "Participant (P001)"
   
   # LANGUAGE SETTINGS
   language_code = "en"  # English
   
   # INTERVIEW CONTEXT DESCRIPTION
   interview_context = (
       "This is an academic research interview about workplace satisfaction. "
       "The interview focuses on employee experiences and organizational culture. "
       "It is conducted between a researcher and a company employee. "
       "The participant works in a mid-sized technology company."
   )
   ```

4. **Run the script:**
   ```bash
   python interview_transcription.py
   ```

## Configuration Guide

### Essential Settings

| Setting | Description | Example |
|---------|-------------|---------|
| `audio_file_path` | Full path to your audio file | `"/Users/john/Desktop/interview.mp3"` |
| `interviewer_name` | How to identify the interviewer | `"Interviewer (Dr. Jane Smith)"` |
| `participant_name` | How to identify the participant | `"Participant (Study ID: P001)"` |
| `language_code` | Two-letter language code | `"en"`, `"da"`, `"de"`, `"fr"` |

### Language Codes

| Language | Code | Language | Code |
|----------|------|----------|------|
| English | `en` | German | `de` |
| Danish | `da` | French | `fr` |
| Spanish | `es` | Italian | `it` |
| Portuguese | `pt` | Dutch | `nl` |
| Swedish | `sv` | Norwegian | `no` |
| Russian | `ru` | Chinese | `zh` |
| Japanese | `ja` | Korean | `ko` |

### Interview Context

The `interview_context` description is crucial for transcription quality. Include:

- **Type of interview**: Research, clinical, journalistic, etc.
- **Main topic/domain**: What the interview is about
- **Participant background**: Role, expertise, context
- **Setting information**: Formal/informal, location context
- **Domain-specific terminology**: Technical terms that might appear

**Good Example:**
```python
interview_context = (
    "This is a clinical psychology research interview about anxiety disorders. "
    "The interview explores patient experiences with cognitive behavioral therapy. "
    "It is conducted between a licensed psychologist and a patient participant. "
    "The participant has been receiving treatment for generalized anxiety disorder. "
    "The discussion includes clinical terminology and treatment methodologies."
)
```

## Output Files

The script creates a timestamped project folder on your desktop containing:

- **`academic_transcript.docx`**: Professional Word document with:
  - Study information table
  - Transcription notes and methodology
  - Detailed statistics
  - Formatted transcript with turn numbers and timestamps

- **`academic_transcript.txt`**: Plain text version of the transcript

- **Audio file copy**: Backup of your original audio file

- **Console output**: Real-time transcript display during processing

## Understanding the Output

### Quality Metrics

- **Audio Quality**: High/Medium/Low based on Whisper's confidence
- **Average Confidence**: Scale from -3 to 0 (higher is better)
- **High Confidence Segments**: Percentage of segments with confidence > -0.5

### Statistics Provided

- Turn-by-turn analysis with timestamps
- Word counts and speech rates
- Pause detection and analysis
- Speaker distribution statistics
- Segment length analysis

### Content Markers

- `[PAUSE]`: Detected significant pauses
- `[UNCLEAR]`: Low-confidence or very short segments

## Troubleshooting

### Common Issues

1. **"ModuleNotFoundError"**: Install missing packages with pip
2. **"FileNotFoundError"**: Check your audio file path is correct
3. **Long processing times**: Normal for large files; Large-v2 model is thorough
4. **Low quality transcription**: Ensure good audio quality and accurate language code

### Performance Tips

- **Audio Quality**: Use high-quality recordings (clear speech, minimal background noise)
- **File Format**: WAV and FLAC generally work best
- **Context Description**: Provide detailed, accurate context for better results
- **Processing Time**: Expect ~1-3x real-time (30-min audio = 30-90 min processing)

### Memory Requirements

- **RAM**: Minimum 4GB available, 8GB+ recommended
- **Storage**: 3GB for Whisper model + space for output files
- **GPU**: Optional but significantly speeds up processing if available

## Advanced Configuration

### Whisper Parameters

The script uses optimized parameters, but you can adjust in the `whisper_params` dictionary:

```python
whisper_params = {
    "temperature": 0.0,  # Lower = more deterministic
    "best_of": 5,        # Number of decoding attempts
    "beam_size": 5,      # Beam search width
    # ... other parameters
}
```

### Speaker Detection

Adjust speaker change detection sensitivity:

```python
speakers = detect_speaker_changes(
    segments, 
    silence_threshold=1.2,  # Seconds of silence to trigger speaker change
    min_speaker_time=4.0    # Minimum time before allowing speaker switch
)
```

## Limitations

- **Speaker Identification**: Based on pauses, not voice recognition
- **Background Noise**: May affect transcription quality
- **Overlapping Speech**: Whisper handles poorly; clean audio recommended
- **Processing Time**: Large files require significant processing time
- **Manual Review**: Always recommended for critical applications

## Support

This tool has been tested on macOS. Windows functionality is theoretical based on standard Python practices. 

For issues:
1. Check your Python and pip installations
2. Verify all dependencies are installed
3. Ensure audio file format is supported
4. Test with a short audio sample first

## License

This tool uses OpenAI's Whisper model. Please review OpenAI's usage terms and any applicable licensing for your research context.

---

**Note**: This tool provides automated transcription with enhanced quality settings. For critical applications, manual review and verification of the transcript is always recommended.