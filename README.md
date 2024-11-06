# Text-to-Speech Synthesis Project

## Overview
This project converts text (or content from `.docx` files) into synthesized speech using Google Cloud Text-to-Speech API, allowing you to generate audio files or play the audio directly. The synthesis process supports French voices, speaking rate customization, and rotating between predefined voices.

## Features
- Converts text to speech using Google Cloud Text-to-Speech API.
- Supports text extraction and preprocessing from `.docx` files.
- Allows customizing speaking rate and voice selection.
- Saves generated audio as `.mp3` files or plays it directly.
- Rotates through a predefined list of French voices.

## Setup

### Prerequisites
- Python 3.11 or later
- A Google Cloud Platform (GCP) account with access to Text-to-Speech API.
- A Google Cloud service account key in JSON format.

### Installation

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd <repository-folder>
