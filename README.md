# Copbot - A 24/7 Offline Emergency Assistant Bot

Copbot is an intelligent, offline chatbot with a graphical user interface designed to provide instant answers to frequently asked questions. It uses advanced NLP techniques including fuzzy matching, TF-IDF vectorization, and FAISS semantic search to deliver accurate responses from an Excel-based knowledge base.

## Features

- **Offline Operation**: Works completely offline - no internet required
- **Multi-layer Search**: Combines fuzzy matching, TF-IDF, and FAISS semantic search for accurate answers
- **User-Friendly GUI**: Clean, modern Tkinter interface with dark theme
- **Voice Assistant Integration**: Quick access to Windows voice assistant (Win+H)
- **Smart Text Processing**: Uses spaCy for lemmatization and stopword removal
- **Excel-Based Knowledge Base**: Easy to update FAQ database using Excel sheets

## Technology Stack

- **GUI Framework**: Tkinter
- **NLP Libraries**: 
  - spaCy (text preprocessing)
  - Sentence Transformers (semantic embeddings)
  - TextBlob (spell correction)
- **Search & Matching**:
  - FAISS (vector similarity search)
  - scikit-learn (TF-IDF vectorization)
  - FuzzyWuzzy (fuzzy string matching)
- **Data Processing**: pandas, numpy

## Prerequisites

- Python 3.8 or higher
- Windows OS (for voice assistant integration)

## Installation

### 1. Install Required Python Packages

```bash
pip install pandas numpy spacy sentence-transformers faiss-cpu scikit-learn textblob fuzzywuzzy python-Levenshtein keyboard
```

### 2. Download spaCy Language Model

```bash
python -m spacy download en_core_web_sm
```

## Configuration

Before running the bot, update the Excel file path in `copfinal.py`:

```python
EXCEL_FILE = r"C:\Users\YourUsername\Path\To\Your\Excel\File.xlsx"
```

### Excel File Format

Your Excel file should contain one or more sheets with the following columns:

- **Question**: The FAQ question text
- **Answer**: The response to provide
- **Additional Info** (optional): Extra context or information

Example:

| Question | Answer | Additional Info |
|----------|--------|-----------------|
| What are the office hours? | Our office is open 9 AM to 5 PM, Monday to Friday. | Closed on public holidays |
| How do I reset my password? | Click on 'Forgot Password' on the login page. | Check your email for reset link |

## Usage

### Running the Application

```bash
python copfinal.py
```

### Using the Chatbot

1. **Type your question** in the input field at the bottom
2. **Press Enter** or click the **Send** button
3. **Receive instant answers** from the knowledge base
4. **Use Voice** button to trigger Windows voice assistant (Win+H)

### Keyboard Shortcuts

- **Enter**: Send message
- **Win+H**: Voice assistant (via Voice button)

## How It Works

1. **Data Loading**: Reads FAQ data from Excel sheets and creates searchable indices
2. **Text Preprocessing**: Normalizes text using lemmatization and removes stopwords
3. **Multi-Stage Matching**:
   - **Fuzzy Matching**: Quick similarity check (threshold: 60%)
   - **TF-IDF**: Keyword-based matching (threshold: 0.35)
   - **FAISS Semantic Search**: Deep semantic understanding (threshold: 0.6)
4. **Response Delivery**: Returns the best matching answer or asks for rephrasing

## File Structure

```
copbot/
├── copfinal.py              # Main application file
├── copfinal_fixed.py        # Improved version with enhanced matching
├── README.md                # This file
└── exe_files/               # Compiled executable artifacts
```

## Troubleshooting

### FAISS Installation Issues
If `faiss-cpu` fails to install, try:
```bash
pip install faiss-cpu --no-cache-dir
```

### FuzzyWuzzy Performance
For faster string matching, install the optional Levenshtein library:
```bash
pip install python-Levenshtein
```

### Voice Assistant Not Working
The keyboard library may require administrator privileges. Run your terminal or IDE as administrator if the voice button doesn't work.

### Excel File Not Found
Ensure the `EXCEL_FILE` path in `copfinal.py` is correct and the file exists at that location.

## Customization

### Adjusting Search Thresholds

In `find_best_answer()` function, you can adjust:
- **Fuzzy matching threshold**: Change `score > 60` (line ~68)
- **TF-IDF threshold**: Change `> 0.35` (line ~79)
- **FAISS threshold**: Change `> 0.6` (line ~86)

### Changing GUI Theme

Modify color values in `CopBotGUI.__init__()`:
- Background: `#1e1e2f`
- Chat frame: `#2e2e42`
- Chat area: `#f4f4f8`
- Buttons: `#4a90e2`

## Future Enhancements

- [ ] Add conversation history export
- [ ] Implement multi-language support
- [ ] Add real-time learning from user feedback
- [ ] Create web-based version
- [ ] Add voice input/output capabilities

## License

This project is provided as-is for educational and commercial use. Feel free to modify and distribute.

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Verify all dependencies are installed correctly
3. Ensure your Excel file follows the required format

## Version History

- **v1.0**: Initial release with fuzzy matching, TF-IDF, and FAISS search
- **v1.1**: Fixed version with improved matching thresholds (copfinal_fixed.py)

---

**Made with ❤️ for 24/7 Emergency Assistance**
