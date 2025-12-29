import tkinter as tk
from tkinter import scrolledtext, messagebox, Menu
import pandas as pd
import faiss
import numpy as np
import spacy
from sentence_transformers import SentenceTransformer
from sklearn.feature_extraction.text import TfidfVectorizer
from textblob import TextBlob
import keyboard
import os
from fuzzywuzzy import process

# Load NLP tools
nlp = spacy.load("en_core_web_sm")
model = SentenceTransformer('all-MiniLM-L6-v2')
vectorizer = TfidfVectorizer()

# Set relative file path
EXCEL_FILE = r"C:\Users\DeLL\OneDrive\Desktop\Cop\College_chatbot(2).xlsx"

def preprocess_text(text):
    """Normalize text using lemmatization and stopword removal."""
    if pd.isna(text):
        return ""
    return " ".join([token.lemma_ for token in nlp(text.lower()) if not token.is_stop])

def load_dataset_and_create_index(excel_file):
    """Load dataset from Excel and create TF-IDF & FAISS indices."""
    try:
        sheets = pd.read_excel(excel_file, sheet_name=None)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load dataset: {e}")
        return [], None, None, None

    faq_list, questions = [], []
    
    for sheet_name, df in sheets.items():
        for _, row in df.iterrows():
            question = preprocess_text(str(row.get("Question", "")))
            answer = str(row.get("Answer", "No answer available."))
            additional_info = str(row.get("Additional Info", ""))
            
            if question.strip():  
                faq_list.append({"category": sheet_name, "question": question, "answer": answer, "additional_info": additional_info})
                questions.append(question)

    if not questions:
        return faq_list, None, None, None  # Avoid crash if dataset is empty

    # Create TF-IDF vectorizer
    tfidf_matrix = vectorizer.fit_transform(questions)

    # FAISS Index for semantic search
    question_vectors = model.encode(questions, convert_to_numpy=True)
    index = faiss.IndexFlatIP(question_vectors.shape[1])
    index.add(question_vectors)
    
    return faq_list, index, question_vectors, tfidf_matrix

def correct_spelling(query):
    """Avoid unnecessary spelling corrections that distort input."""
    return str(TextBlob(query).correct()) if len(query.split()) <= 3 else query

def find_best_answer(user_query, faq_list, index=None, tfidf_matrix=None):
    user_query = user_query.lower()  # Normalize input

    # Extract best match using fuzzy search
    questions = [item["question"] for item in faq_list]
    best_match, score = process.extractOne(user_query, questions)

    if score > 60:  # Adjust threshold as needed
        for item in faq_list:
            if item["question"] == best_match:
                return item


    return {"answer": "I'm not sure. Please try rephrasing."}


    # TF-IDF Matching
    query_tfidf = vectorizer.transform([user_query])
    keyword_scores = (tfidf_matrix @ query_tfidf.T).toarray().flatten()
    keyword_idx = np.argmax(keyword_scores)

    if keyword_scores[keyword_idx] > 0.35:  
        return faq_list[keyword_idx]
    
    # FAISS Semantic Search
    query_vector = model.encode([user_query], convert_to_numpy=True)
    scores, best_match_idx = index.search(query_vector, 1)

    if scores[0][0] > 0.6:  
        return faq_list[best_match_idx[0][0]]

    return {"answer": "I'm not sure. Please try rephrasing."}

class CopBotGUI:
    def __init__(self, root, faq_list, index, tfidf_matrix):
        self.root = root
        self.faq_list = faq_list
        self.index = index
        self.tfidf_matrix = tfidf_matrix
        self.root.title("CopBotChatbox - Offline")
        self.root.geometry("750x650")
        self.root.configure(bg="#1e1e2f")
        self.create_menu()
        
        tk.Label(self.root, text="CopBotChatbox", font=("Helvetica", 24, "bold"), bg="#1e1e2f", fg="#e8e8e8").pack(pady=(20, 10))
        
        self.chat_frame = tk.Frame(self.root, bg="#2e2e42", bd=2, relief=tk.GROOVE)
        self.chat_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        self.chat_area = scrolledtext.ScrolledText(self.chat_frame, wrap=tk.WORD, state='disabled', font=("Helvetica", 12), bg="#f4f4f8", fg="#1e1e2f", relief=tk.FLAT, padx=15, pady=15)
        self.chat_area.pack(fill=tk.BOTH, expand=True)
        
        self.input_frame = tk.Frame(self.root, bg="#1e1e2f")
        self.input_frame.pack(padx=20, pady=(10, 20), fill=tk.X)
        
        self.user_input = tk.Entry(self.input_frame, font=("Helvetica", 12), bd=2, relief=tk.GROOVE)
        self.user_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.user_input.bind("<Return>", self.send_message)
        
        self.button_frame = tk.Frame(self.input_frame, bg="#1e1e2f")
        self.button_frame.pack(side=tk.RIGHT)
        
        self.voice_button = tk.Button(self.button_frame, text="Voice", font=("Helvetica", 12, "bold"), bg="#4a90e2", fg="white", bd=0, padx=15, pady=8, command=self.trigger_voice_assistant, activebackground="#357ab8", cursor="hand2")
        self.voice_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.send_button = tk.Button(self.button_frame, text="Send", font=("Helvetica", 12, "bold"), bg="#4a90e2", fg="white", bd=0, padx=20, pady=8, command=self.send_message, activebackground="#357ab8", cursor="hand2")
        self.send_button.pack(side=tk.LEFT)
        
        self.display_message("CopBot", "Welcome! How can I assist you today?")
    
    def create_menu(self):
        """Create a menu bar."""
        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)
        file_menu = Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
    
    def display_message(self, sender, message):
        """Display chat messages."""
        self.chat_area.config(state='normal')
        self.chat_area.insert(tk.END, f"{sender}: {message}\n\n")
        self.chat_area.config(state='disabled')
        self.chat_area.see(tk.END)
    
    def send_message(self, event=None):
        """Handle user input and fetch response."""
        message = self.user_input.get().strip()
        if not message:
            return
        self.display_message("You", message)
        self.user_input.delete(0, tk.END)
        result = find_best_answer(message, self.faq_list, self.index, self.tfidf_matrix)
        self.display_message("CopBot", result["answer"])
    
    def trigger_voice_assistant(self):
        """Trigger Windows voice assistant."""
        try:
            keyboard.press_and_release('win+h')
        except Exception as e:
            messagebox.showerror("Error", f"Unable to trigger voice assistant: {e}")

def main():
    """Main function to run the chatbot."""
    faq_list, index, _, tfidf_matrix = load_dataset_and_create_index(EXCEL_FILE)
    root = tk.Tk()
    app = CopBotGUI(root, faq_list, index, tfidf_matrix)
    root.mainloop()

if __name__ == "__main__":
    main()
