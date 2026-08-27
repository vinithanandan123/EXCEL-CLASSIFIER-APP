import tkinter as tk
from tkinter import filedialog, Label, Button, messagebox, scrolledtext, ttk, Toplevel, Entry, Menu, simpledialog
from PIL import Image, ImageTk
import numpy as np
import os
import cv2
import openpyxl
from openpyxl import Workbook
import time
import re
import random
import pyttsx3
import webbrowser
import speech_recognition as sr
import subprocess
import ctypes
import sys
import threading
from datetime import datetime, timedelta
import pandas as pd
import requests
import csv
import pyautogui  # For screenshot functionality
import wikipedia  # Added missing import

# --- Configuration Constants ---
# JARVIS Theme Colors
BG_DARK = "#1C2526"
BG_MEDIUM = "#2D3839"
TEXT_PRIMARY = "#00DDEB"  # JARVIS Blue
TEXT_SECONDARY = "#C4C4C4"  # Grayish White
ACCENT_RED = "#FF6F61"  # Fire Red / Alert Red

# Base directories for memory and models (created relative to script's execution directory)
MEMORY_BASE_DIR = "memory"
MODEL_BASE_DIR = "model_memory"


# --- Activation Functions for Neural Network ---
def sigmoid(x):
    """Sigmoid activation function."""
    # Added a check to prevent overflow in np.exp
    return 1 / (1 + np.exp(-np.clip(x, -500, 500)))


def sigmoid_derivative(x):
    """Derivative of the sigmoid function."""
    return x * (1 - x)


# --- Neural Network Class (for Fire Detection) ---
class SimpleNN:
    def __init__(self, input_size):
        """
        Initializes a simple two-layer neural network.
        :param input_size: Number of input features (e.g., 64*64 for a 64x64 image).
        """
        self.input_size = input_size
        # Weights and biases for the hidden layer (64 neurons)
        self.weights1 = np.random.randn(input_size, 64)
        self.bias1 = np.zeros((1, 64))
        # Weights and biases for the output layer (1 neuron for binary classification)
        self.weights2 = np.random.randn(64, 1)
        self.bias2 = np.zeros((1, 1))

    def forward(self, X):
        """
        Performs a forward pass through the network.
        :param X: Input data (numpy array).
        :return: Output of the network.
        """
        self.z1 = np.dot(X, self.weights1) + self.bias1
        self.a1 = sigmoid(self.z1)  # Activation of hidden layer
        self.z2 = np.dot(self.a1, self.weights2) + self.bias2
        self.output = sigmoid(self.z2)  # Activation of output layer
        return self.output

    def train(self, X, y, epochs=500, lr=0.1, progress_callback=None):
        """
        Trains the neural network using backpropagation.
        :param X: Training input data.
        :param y: True labels for training data.
        :param epochs: Number of training iterations.
        :param lr: Learning rate.
        :param progress_callback: Optional function to update training progress in GUI.
        """
        for epoch in range(epochs):
            self.forward(X)  # Forward pass

            # Calculate output layer error and delta
            error = y - self.output
            d_output = error * sigmoid_derivative(self.output)

            # Calculate hidden layer error and delta
            error_hidden = d_output.dot(self.weights2.T)
            d_hidden = error_hidden * sigmoid_derivative(self.a1)

            # Update weights and biases
            self.weights2 += self.a1.T.dot(d_output) * lr
            self.bias2 += np.sum(d_output, axis=0, keepdims=True) * lr
            self.weights1 += X.T.dot(d_hidden) * lr
            self.bias1 += np.sum(d_hidden, axis=0, keepdims=True) * lr

            if progress_callback:
                progress_callback(epoch, epochs)
        return self.output

    def predict(self, X):
        """
        Makes a prediction using the trained network.
        :param X: Input data for prediction.
        :return: Predicted output.
        """
        return self.forward(X)

    def save(self, folder):
        """
        Saves the network's weights and biases to specified folder.
        :param folder: Directory to save the model files.
        """
        os.makedirs(folder, exist_ok=True)  # Create folder if it doesn't exist
        np.save(os.path.join(folder, "weights1.npy"), self.weights1)
        np.save(os.path.join(folder, "bias1.npy"), self.bias1)
        np.save(os.path.join(folder, "weights2.npy"), self.weights2)
        np.save(os.path.join(folder, "bias2.npy"), self.bias2)

    @classmethod
    def load(cls, input_size, folder):
        """
        Loads a saved neural network model from a folder.
        :param input_size: Input size of the model.
        :param folder: Directory where model files are saved.
        :return: Loaded SimpleNN instance.
        :raises FileNotFoundError: If model files are missing.
        """
        # Check if all necessary model files exist before attempting to load
        required_files = ["weights1.npy", "bias1.npy", "weights2.npy", "bias2.npy"]
        if not all(os.path.exists(os.path.join(folder, f)) for f in required_files):
            raise FileNotFoundError(f"One or more model files not found in {folder}. Please train the model first.")

        model = cls(input_size)
        model.weights1 = np.load(os.path.join(folder, "weights1.npy"))
        model.bias1 = np.load(os.path.join(folder, "bias1.npy"))
        model.weights2 = np.load(os.path.join(folder, "weights2.npy"))
        model.bias2 = np.load(os.path.join(folder, "bias2.npy"))
        return model


# --- Image Preprocessing Functions (for Fire Detection) ---
def load_image(path):
    """
    Loads and preprocesses an image for the neural network.
    Converts to grayscale, resizes to 64x64, flattens, and normalizes.
    :param path: Path to the image file.
    :return: Flattened and normalized image vector, or None if invalid.
    """
    img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
    if img is None:
        messagebox.showerror("Image Error", "Invalid image file. Please select a valid image (jpg, png, jpeg).")
        return None
    img = cv2.resize(img, (64, 64))  # Resize to 64x64 pixels
    return img.flatten() / 255.0  # Flatten to a 1D array and normalize pixel values


def load_dataset(folder):
    """
    Loads images from a dataset folder structured with 'fire' and 'nofire' subfolders.
    :param folder: Base directory of the dataset.
    :return: Tuple of (X, y) numpy arrays, where X is image data and y is labels.
    """
    X, y = [], []
    # Iterate through 'nofire' (label 0) and 'fire' (label 1) folders
    for label, class_folder in enumerate(['nofire', 'fire']):
        class_path = os.path.join(folder, class_folder)
        if not os.path.exists(class_path):
            print(f"Warning: Class folder '{class_path}' not found. Skipping.")
            continue  # Skip if folder doesn't exist

        for filename in os.listdir(class_path):
            img_path = os.path.join(class_path, filename)
            try:
                img_vector = load_image(img_path)
                if img_vector is not None:
                    X.append(img_vector)
                    y.append(label)
            except Exception as e:
                print(f"Skipping image {img_path} due to error during loading/processing: {e}")

    if not X:
        return np.array([]), np.array([])  # Return empty arrays if no images were loaded
    return np.array(X), np.array(y).reshape(-1, 1)  # Reshape y to be a column vector


# --- Chat Application Class (JARVIS) ---
class ChatApp(tk.Frame):
    def __init__(self, master, ai_name="JARVIS"):
        """
        Initializes the Chat AI application GUI with a specific AI personality.
        :param master: The parent Tkinter widget (usually a ttk.Frame or Toplevel).
        :param ai_name: The name of the AI personality (e.g., "JARVIS", "FRIDAY").
        """
        super().__init__(master, bg=BG_DARK)
        self.pack(fill="both", expand=True)
        self.master = master  # Store master for messagebox parent
        self.ai_name = ai_name  # Current AI personality name

        # Define memory paths based on AI name
        self.memory_path = os.path.join(MEMORY_BASE_DIR, self.ai_name.lower())
        os.makedirs(self.memory_path, exist_ok=True)  # Ensure AI's specific memory folder exists

        self.memory_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_chat_memory.xlsx")
        self.birthday_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_birthdays.csv")
        self.schedule_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_study_schedule.xlsx")
        self.english_data_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_english_data.xlsx")

        # Default column headers for chat memory Excel
        self.column_headers = {
            "user": "User",
            "input": "Input",
            "response": "Response",
            "name": "Name",
            "birthday": "Birthday",
            "address": "Address",
            "number": "Number"
        }
        self.user_info = {"name": "", "birthday": "", "address": "", "number": ""}  # Stores user personal info
        self.birthdays = {}  # Stores birthdays from CSV
        self.custom_column_count = 0  # Tracks number of custom columns added

        # Jokes specific to the AI (or can be shared)
        self.joke_list = [
            f"{self.ai_name}: Why did the scarecrow become a motivational speaker? Because he was outstanding in his field!",
            f"{self.ai_name}: Why can't basketball players go on vacation? Because they would get called for traveling.",
            f"{self.ai_name}: What do you call a bear with no socks on? Barefoot!"
        ]

        # Text-to-Speech (TTS) engine initialization
        self.tts_engine = pyttsx3.init()
        self.tts_engine.setProperty('rate', 150)  # Speed of speech
        self.tts_engine.setProperty('volume', 1)  # Max volume
        try:
            # Attempt to set a specific male voice for Windows
            self.tts_engine.setProperty('voice',
                                        'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_DAVID_11.0')
        except Exception:
            print("David voice not found, using default voice for TTS.")

        # Speech Recognition (SR) setup
        self.recognizer = sr.Recognizer()
        self.listening = False  # Flag to control continuous listening
        self.listen_thread = None  # Thread for continuous listening

        # DataFrames for study schedule and English lessons
        self.df_marks = None  # Stores marks data
        self.df_schedule_data = None  # Stores schedule data
        self.schedule = {}  # Dictionary for quick schedule lookup
        self.english_levels = {}  # Dictionary for English teaching data

        self.setup_ui()  # Build the GUI
        self.prepare_excel_files()  # Ensure all necessary Excel/CSV files exist
        self.load_all_data()  # Load all data from files

    def speak_async(self, text):
        """Speaks the given text asynchronously to avoid blocking the GUI."""
        threading.Thread(target=lambda: (self.tts_engine.say(text), self.tts_engine.runAndWait()), daemon=True).start()

    def display_async(self, text, tag=None):
        """
        Displays text in the chat history asynchronously.
        :param text: The text to display.
        :param tag: Optional tag for text styling (e.g., "user", "ai").
        """

        def run_display():
            self.chat_display.config(state='normal')  # Enable editing
            self.chat_display.insert(tk.END, text + "\n", tag)  # Insert text
            self.chat_display.yview(tk.END)  # Scroll to the end
            self.chat_display.config(state='disabled')  # Disable editing

        self.master.after(0, run_display)  # Use master.after to safely update GUI from any thread

    def setup_ui(self):
        """Sets up all the GUI elements for the Chat AI tab."""
        # Configure ttk styles for consistent theming
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TProgressbar", troughcolor=BG_MEDIUM, background=TEXT_PRIMARY, thickness=20)
        style.configure("Jarvis.TButton", background=BG_MEDIUM, foreground=TEXT_PRIMARY,
                        font=("Arial", 10, "bold"), relief="flat", borderwidth=0)
        style.map("Jarvis.TButton",
                  background=[('active', BG_MEDIUM)],
                  foreground=[('active', TEXT_PRIMARY)])

        # Top section for AI Name and Time/Date display
        top_frame = tk.Frame(self, bg=BG_DARK)
        top_frame.pack(fill="x", pady=5)

        Label(top_frame, text=f"Active AI: {self.ai_name}", font=("Arial", 16, "bold"), fg=TEXT_PRIMARY,
              bg=BG_DARK).pack(side="left", padx=10)

        self.time_label = Label(top_frame, font=("Helvetica", 16), fg=TEXT_PRIMARY, bg=BG_DARK)
        self.time_label.pack(side="right", padx=10)
        self.update_time()  # Start time update loop

        self.date_label = Label(top_frame, font=("Helvetica", 12), fg=TEXT_SECONDARY, bg=BG_DARK)
        self.date_label.pack(side="right", padx=10)
        self.update_date()  # Start date update loop

        # Chat History Display Area
        self.chat_display = scrolledtext.ScrolledText(self, wrap=tk.WORD, width=60, height=20, state='disabled',
                                                      bg=BG_MEDIUM, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
                                                      font=("Arial", 10))
        self.chat_display.pack(pady=10, fill="both", expand=True)
        # Configure tags for user and AI messages for different colors
        self.chat_display.tag_configure("user", foreground=TEXT_SECONDARY)
        self.chat_display.tag_configure("ai", foreground=TEXT_PRIMARY)
        # Initial greeting from the AI
        self.chat_display.insert(tk.END, f"{self.ai_name}: Hey, I’m {self.ai_name}—ready to help you out!\n", "ai")

        # Input Field and Send Button
        input_frame = tk.Frame(self, bg=BG_DARK)
        input_frame.pack(fill="x", pady=5)
        self.entry = tk.Entry(input_frame, width=50, bg=BG_MEDIUM, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
                              font=("Arial", 10))
        self.entry.pack(side=tk.LEFT, padx=10, fill="x", expand=True)
        self.entry.bind("<Return>", lambda event: self.handle_send())  # Bind Enter key to send message

        self.send_button = Button(input_frame, text="Send", command=self.handle_send,
                                  bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 10), activebackground=BG_MEDIUM)
        self.send_button.pack(side=tk.LEFT, padx=5)

        # Control Buttons (Voice, Memory, Edit Headers, Train)
        control_frame = tk.Frame(self, bg=BG_DARK)
        control_frame.pack(fill="x", pady=5)

        buttons_left = [
            ("🎤 Start Listen", self.start_listening),
            ("⏹ Stop Listen", self.stop_listening),
            ("📂 Select Memory Folder", self.select_memory_folder),
            ("✏️ Edit Column Headers", self.edit_column_headers),
            (f"⚙️ Train {self.ai_name} AI", self.train_chat)  # Button text updates with AI name
        ]

        # Use a grid layout for these buttons for better organization
        for i, (text, command) in enumerate(buttons_left):
            Button(control_frame, text=text, command=command,
                   bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 9), activebackground=BG_MEDIUM).grid(row=i // 3,
                                                                                                      column=i % 3,
                                                                                                      padx=5, pady=2,
                                                                                                      sticky="ew")

        # Progress bar and status for memory processing/training
        self.progress_bar = ttk.Progressbar(self, length=300, mode='determinate', style="TProgressbar")
        self.progress_bar.pack(pady=5)
        self.progress_bar.pack_forget()  # Hide initially

        self.loading_indicator = Label(self, text="", font=("Arial", 10), fg=TEXT_PRIMARY, bg=BG_DARK)
        self.loading_indicator.pack(pady=5)
        self.loading_indicator.pack_forget()  # Hide initially

        self.memory_status = Label(self, text=f"Memory folder: {self.memory_path}", fg=TEXT_SECONDARY, bg=BG_DARK,
                                   font=("Arial", 10))
        self.memory_status.pack(pady=5)

        # Media and Utility Buttons (bottom section)
        utility_frame = tk.Frame(self, bg=BG_DARK)
        utility_frame.pack(fill="x", pady=5)

        utility_buttons = [
            ("Open Notepad", self.open_notepad), ("Play Music", self.play_music),
            ("Open YouTube", self.open_youtube), ("Google Search", self.google_search),
            ("Take Screenshot", self.take_screenshot), ("Open Spotify", self.open_spotify),
            ("Open WhatsApp", self.open_whatsapp)
        ]

        # Use a grid layout for utility buttons
        for i, (text, command) in enumerate(utility_buttons):
            Button(utility_frame, text=text, command=command,
                   bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 9), activebackground=BG_MEDIUM).grid(row=i // 4,
                                                                                                      column=i % 4,
                                                                                                      padx=5, pady=2,
                                                                                                      sticky="ew")

        # Button to open the separate mini controls window
        self.mini_controls_window = None  # Initialize to None
        Button(utility_frame, text="Open Controls", command=self.create_mini_window,
               bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 9), activebackground=BG_MEDIUM).grid(row=1, column=3,
                                                                                                  padx=5, pady=2,
                                                                                                  sticky="ew")

    def prepare_excel_files(self):
        """
        Ensures all necessary Excel and CSV files for the AI's memory exist.
        Creates them with initial headers/data if they don't.
        """
        # Prepare chat memory Excel file
        if not os.path.exists(self.memory_file):
            wb = Workbook()
            ws = wb.active
            ws.title = "Chat History"  # Set sheet title
            ws.append([self.column_headers[key] for key in self.column_headers])  # Write headers
            wb.save(self.memory_file)

        # Prepare birthdays CSV file
        if not os.path.exists(self.birthday_file):
            with open(self.birthday_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Birthday"])  # Write header row

        # Prepare study schedule Excel file with two sheets: 'Marks' and 'Schedule'
        if not os.path.exists(self.schedule_file):
            # Initial data for 'Marks' sheet
            initial_marks_data = {
                "Subject": ["Subject 1", "Subject 2", "Subject 3", "Subject 4", "Subject 5", "Subject 6"],
                "Test-01": [0, 0, 0, 0, 0, 0],
                "Test-02": [0, 0, 0, 0, 0, 0],
                "Test-03": [0, 0, 0, 0, 0, 0],
                "Assignment": [0, 0, 0, 0, 0, 0]
            }
            df_marks_initial = pd.DataFrame(initial_marks_data)

            # Initial data for 'Schedule' sheet
            initial_schedule_data = {
                "Date": [],
                "Event": []
            }
            # Add a couple of sample schedule entries for today and tomorrow
            today = datetime.now()
            initial_schedule_data["Date"].append(today.strftime("%b %d").replace(" 0", " "))
            initial_schedule_data["Event"].append("6:00-8:00 AM: Morning Review; 8:00 AM - 3:30 PM: Test at college")

            tomorrow = today + timedelta(days=1)
            initial_schedule_data["Date"].append(tomorrow.strftime("%b %d").replace(" 0", " "))
            initial_schedule_data["Event"].append("Evening: Project Work; 11:30 PM - 6:00 AM: Sleep (6.5h)")

            df_schedule_initial = pd.DataFrame(initial_schedule_data)

            # Save both DataFrames to different sheets in the same Excel file
            with pd.ExcelWriter(self.schedule_file, engine='openpyxl') as writer:
                df_marks_initial.to_excel(writer, sheet_name='Marks', index=False)
                df_schedule_initial.to_excel(writer, sheet_name='Schedule', index=False)

        # Prepare English teaching data Excel file with multiple levels (sheets)
        if not os.path.exists(self.english_data_file):
            wb_english = Workbook()

            # Level 1
            ws_level1 = wb_english.create_sheet("Level 1", 0)  # Create at index 0
            ws_level1.append(["Sentence", "Explanation"])
            ws_level1.append(["Hello, how are you?", "A common greeting."])
            ws_level1.append(["My name is Alex.", "Introducing yourself."])
            ws_level1.append(["I like to read books.", "Stating a preference."])
            ws_level1.append(["It is a sunny day.", "Describing the weather."])
            ws_level1.append(["Can I help you?", "Offering assistance."])

            # Level 2
            ws_level2 = wb_english.create_sheet("Level 2")
            ws_level2.append(["Sentence", "Explanation"])
            ws_level2.append(
                ["The quick brown fox jumps over the lazy dog.", "A pangram, containing all letters of the alphabet."])
            ws_level2.append(["She sings beautifully.", "Describing an action with an adverb."])
            ws_level2.append(["I went to the store yesterday.", "Simple past tense usage."])
            ws_level2.append(["He is studying for his exam.", "Present continuous tense."])
            ws_level2.append(["They will arrive soon.", "Future tense with 'will'."])

            # Level 3 (Example)
            ws_level3 = wb_english.create_sheet("Level 3")
            ws_level3.append(["Sentence", "Explanation"])
            ws_level3.append(["Although it was raining, we went for a walk.", "Using 'although' for contrast."])
            ws_level3.append(["If I had known, I would have told you.", "Third conditional for hypothetical past."])
            ws_level3.append(["The book, which was very old, fell apart.", "Using a relative clause."])
            ws_level3.append(["Despite the challenges, they succeeded.", "Using 'despite' for contrast."])
            ws_level3.append(["She's accustomed to waking up early.", "Using 'accustomed to' for habit."])

            # Remove default 'Sheet' created by Workbook() if it's empty
            if 'Sheet' in wb_english.sheetnames and wb_english['Sheet'].max_row == 1 and wb_english[
                'Sheet'].max_column == 1:
                del wb_english['Sheet']

            wb_english.save(self.english_data_file)

    def load_all_data(self):
        """
        Loads all persistent data for the current AI personality:
        chat history, user info, birthdays, study schedule, and English lessons.
        """
        # Load chat user info from chat_memory.xlsx
        if os.path.exists(self.memory_file):
            try:
                wb = openpyxl.load_workbook(self.memory_file)
                ws = wb.active
                # Read header row to get column indices dynamically
                header_row = [cell.value for cell in ws[1]]

                # Update column_headers dictionary based on loaded headers,
                # and count custom columns
                self.custom_column_count = 0
                new_column_headers = {}
                for i, col_name in enumerate(header_row):
                    if col_name in self.column_headers.values():  # Check if it's a known header
                        # Find the key for this value
                        for k, v in self.column_headers.items():
                            if v == col_name:
                                new_column_headers[k] = col_name
                                break
                    else:  # Assume it's a custom column
                        self.custom_column_count += 1
                        new_column_headers[f"custom{self.custom_column_count}"] = col_name
                self.column_headers = new_column_headers

                # Load user_info from the last row (most recent data)
                if ws.max_row > 1:  # Check if there's data beyond headers
                    last_row_values = list(ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, values_only=True))[0]

                    # Map values to user_info using dynamic column indices
                    for key, header_name in self.column_headers.items():
                        if key in ["name", "birthday", "address", "number"]:  # Only update these specific fields
                            try:
                                col_idx = header_row.index(header_name)
                                if col_idx < len(last_row_values) and last_row_values[col_idx] is not None:
                                    self.user_info[key] = str(last_row_values[col_idx]).strip()
                            except ValueError:
                                pass  # Column not found, skip

            except Exception as e:
                self.memory_status.config(text=f"❌ Error loading chat memory: {str(e)}", fg=ACCENT_RED)

        # Load birthdays from CSV
        if os.path.exists(self.birthday_file):
            try:
                df_birth = pd.read_csv(self.birthday_file)
                self.birthdays = dict(zip(df_birth["Name"], df_birth["Birthday"]))
            except Exception as e:
                print(f"Error loading birthday file {self.birthday_file}: {e}")
                self.birthdays = {}

        # Load study schedule from Excel (Marks and Schedule sheets)
        if os.path.exists(self.schedule_file):
            try:
                self.df_marks = pd.read_excel(self.schedule_file, sheet_name='Marks')
                self.df_schedule_data = pd.read_excel(self.schedule_file, sheet_name='Schedule')
                # Convert schedule DataFrame to a dict for easier lookup
                self.schedule = dict(zip(self.df_schedule_data["Date"], self.df_schedule_data["Event"]))
            except Exception as e:
                print(f"Error loading study schedule file {self.schedule_file}: {e}")
                self.df_marks = pd.DataFrame()  # Reset to empty if error
                self.schedule = {}
        else:
            self.df_marks = pd.DataFrame()  # No file, so empty
            self.schedule = {}

        # Load English teaching data from Excel
        self.english_levels = {}
        if os.path.exists(self.english_data_file):
            try:
                wb_english = openpyxl.load_workbook(self.english_data_file)
                for sheet_name in wb_english.sheetnames:
                    if sheet_name.startswith("Level "):  # Only process sheets named "Level X"
                        df_level = pd.read_excel(self.english_data_file, sheet_name=sheet_name)
                        if 'Sentence' in df_level.columns and 'Explanation' in df_level.columns:
                            # Store as a list of (sentence, explanation) tuples
                            self.english_levels[sheet_name] = list(zip(df_level['Sentence'], df_level['Explanation']))
            except Exception as e:
                print(f"Error loading English data file {self.english_data_file}: {e}")
                self.english_levels = {}  # Reset to empty if error

    def select_memory_folder(self):
        """Allows user to select a new base folder for AI memory files."""
        new_memory_base_path = filedialog.askdirectory(parent=self.master)
        if new_memory_base_path:
            # Update the AI's specific memory path
            self.memory_path = os.path.join(new_memory_base_path, self.ai_name.lower())
            os.makedirs(self.memory_path, exist_ok=True)  # Create the new AI-specific folder

            # Update all file paths
            self.memory_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_chat_memory.xlsx")
            self.birthday_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_birthdays.csv")
            self.schedule_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_study_schedule.xlsx")
            self.english_data_file = os.path.join(self.memory_path, f"{self.ai_name.lower()}_english_data.xlsx")

            self.memory_status.config(text=f"Memory folder: {self.memory_path}", fg=TEXT_PRIMARY)
            self.prepare_excel_files()  # Re-prepare files in the new location (creates if not existing)
            self.load_all_data()  # Reload all data from the new location

    def edit_column_headers(self):
        """Opens a Toplevel window to allow editing of Excel column headers."""
        window = Toplevel(self.master, bg=BG_DARK)
        window.title("Edit Column Headers")

        entries = {}
        # Fixed headers (cannot be removed, but names can be changed)
        fixed_headers_keys = ["user", "input", "response", "name", "birthday", "address", "number"]
        for i, key in enumerate(fixed_headers_keys):
            Label(window, text=f"Header for {key.capitalize()}:", fg=TEXT_PRIMARY, bg=BG_DARK, font=("Arial", 10)).grid(
                row=i, column=0, padx=5, pady=5)
            entries[key] = Entry(window, bg=BG_MEDIUM, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
                                 font=("Arial", 10))
            entries[key].insert(0, self.column_headers[key])  # Display current header
            entries[key].grid(row=i, column=1, padx=5, pady=5)

        # Custom headers (can be added dynamically)
        current_row_for_custom = len(fixed_headers_keys)
        # Populate existing custom headers if any
        for i in range(self.custom_column_count):
            key = f"custom{i + 1}"
            Label(window, text=f"Header for Custom {i + 1}:", fg=TEXT_PRIMARY, bg=BG_DARK, font=("Arial", 10)).grid(
                row=current_row_for_custom + i, column=0, padx=5, pady=5)
            entries[key] = Entry(window, bg=BG_MEDIUM, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
                                 font=("Arial", 10))
            entries[key].insert(0, self.column_headers.get(key, f"Custom{i + 1}"))  # Get existing or default
            entries[key].grid(row=current_row_for_custom + i, column=1, padx=5, pady=5)

        def add_column():
            """Adds a new custom column entry field to the Toplevel window."""
            self.custom_column_count += 1
            new_key = f"custom{self.custom_column_count}"
            self.column_headers[new_key] = f"Custom{self.custom_column_count}"  # Add to current headers for display
            Label(window, text=f"Header for Custom {self.custom_column_count}:", fg=TEXT_PRIMARY, bg=BG_DARK,
                  font=("Arial", 10)).grid(row=current_row_for_custom + self.custom_column_count - 1, column=0, padx=5,
                                           pady=5)
            entries[new_key] = Entry(window, bg=BG_MEDIUM, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
                                     font=("Arial", 10))
            entries[new_key].insert(0, self.column_headers[new_key])
            entries[new_key].grid(row=current_row_for_custom + self.custom_column_count - 1, column=1, padx=5, pady=5)
            window.update_idletasks()  # Update window geometry to fit new elements
            window.geometry(f"300x{window.winfo_height()}")

        def save_headers():
            """Saves the updated column headers and prepares the Excel file."""
            new_column_headers = {}
            # Process fixed headers
            for key in fixed_headers_keys:
                new_header = entries[key].get().strip()
                if new_header:
                    new_column_headers[key] = new_header
                else:
                    new_column_headers[key] = self.column_headers[key]  # Keep old if empty

            # Process custom headers
            for i in range(self.custom_column_count):
                key = f"custom{i + 1}"
                new_header = entries[key].get().strip()
                if new_header:
                    new_column_headers[key] = new_header
                else:
                    new_column_headers[key] = self.column_headers.get(key, f"Custom{i + 1}")

            self.column_headers = new_column_headers  # Update the instance's column_headers
            self.prepare_excel_files()  # This will ensure the Excel file's header row is consistent
            messagebox.showinfo("Success", "Column headers updated! New chat entries will use these headers.",
                                parent=window)
            window.destroy()

        Button(window, text="Add New Column", command=add_column,
               bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 10), activebackground=BG_MEDIUM).grid(
            row=current_row_for_custom + self.custom_column_count, column=0, columnspan=2, pady=5)
        Button(window, text="Save", command=save_headers,
               bg=BG_MEDIUM, fg=TEXT_PRIMARY, font=("Arial", 10), activebackground=BG_MEDIUM).grid(
            row=current_row_for_custom + self.custom_column_count + 1, column=0, columnspan=2, pady=5)

    def update_progress_chat(self, step, total_steps):
        """
        Updates the progress bar and loading indicator for chat AI memory processing.
        :param step: Current step number.
        :param total_steps: Total number of steps.
        """
        progress = (step + 1) / total_steps * 100
        self.progress_bar['value'] = progress
        spinner = ["|", "/", "-", "\\"][(step // 5) % 4]
        self.loading_indicator.config(text=f"Processing {spinner}")
        self.master.update_idletasks()  # Force GUI update

    def train_chat(self):
        """
        Initiates the process of loading and "processing" the chat AI's memory.
        This reloads all data from Excel/CSV files.
        """
        if not os.path.exists(self.memory_file):
            self.memory_status.config(text="⚠️ No memory file found! Select a folder or start chatting.", fg=ACCENT_RED)
            return

        self.memory_status.config(text=f"🧠 Processing {self.ai_name} memory...", fg=TEXT_PRIMARY)
        self.progress_bar.pack()
        self.loading_indicator.pack()
        self.progress_bar['value'] = 0
        self.loading_indicator.config(text="Processing |")
        self.master.update_idletasks()

        try:
            self.load_all_data()  # Reload all data (user_info, birthdays, schedule, English data)
            self.memory_status.config(text=f"✅ {self.ai_name} memory processed!", fg=TEXT_PRIMARY)
        except Exception as e:
            self.memory_status.config(text=f"❌ Error processing memory: {str(e)}", fg=ACCENT_RED)
        finally:
            self.progress_bar.pack_forget()
            self.loading_indicator.pack_forget()

    def find_previous_response(self, message):
        """
        Searches the chat memory Excel file for a previous response to an identical message.
        :param message: The user's input message.
        :return: The AI's previous response if found, otherwise None.
        """
        if not os.path.exists(self.memory_file):
            return None
        try:
            wb = openpyxl.load_workbook(self.memory_file)
            ws = wb.active
            header_row = [cell.value for cell in ws[1]]
            try:
                # Get column indices dynamically based on current headers
                input_col_idx = header_row.index(self.column_headers["input"])
                response_col_idx = header_row.index(self.column_headers["response"])
            except ValueError:
                print("Input or Response column not found in memory file headers. Cannot find previous response.")
                return None

            for row in ws.iter_rows(min_row=2, values_only=True):  # Start from second row (after headers)
                # Ensure row has enough columns and input is not None
                if len(row) > input_col_idx and row[input_col_idx] is not None and str(
                        row[input_col_idx]).lower() == message.lower():
                    if len(row) > response_col_idx:
                        return row[response_col_idx]
        except Exception as e:
            print(f"Error finding previous response in Excel: {e}")
        return None

    def store_to_excel(self, user, message, response, **kwargs):
        """
        Stores a conversation entry (user input, AI response, and user info) into the Excel memory file.
        :param user: The user's identifier (e.g., "User").
        :param message: The user's input message.
        :param response: The AI's response.
        :param kwargs: Additional keyword arguments for custom columns.
        """
        try:
            wb = openpyxl.load_workbook(self.memory_file)
            ws = wb.active

            # Ensure header consistency. If headers have changed, this might lead to new columns
            # or data misalignment if not handled carefully. For this app, we append based on
            # the current `self.column_headers` order.

            row_data = []
            for key in self.column_headers:
                if key == "user":
                    row_data.append(user)
                elif key == "input":
                    row_data.append(message)
                elif key == "response":
                    row_data.append(response)
                elif key in self.user_info:  # For fixed user info fields
                    row_data.append(self.user_info.get(key, ""))
                elif key.startswith("custom"):  # For dynamic custom fields
                    row_data.append(kwargs.get(key, ""))
                else:
                    row_data.append("")  # Fallback for any unexpected columns in self.column_headers

            ws.append(row_data)  # Append the new row
            wb.save(self.memory_file)
        except Exception as e:
            print(f"Error storing to Excel: {e}")
            messagebox.showerror("Excel Save Error", f"Could not save chat memory to Excel: {e}", parent=self.master)

    def save_birthdays(self):
        """Saves the current birthday data to the AI's specific CSV file."""
        try:
            with open(self.birthday_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Birthday"])  # Write header
                for name, dob in self.birthdays.items():
                    writer.writerow([name, dob])  # Write each entry
        except Exception as e:
            print(f"Error saving birthdays: {e}")
            messagebox.showerror("Birthday Save Error", f"Could not save birthdays: {e}", parent=self.master)

    def check_birthdays(self):
        """Checks if any birthdays match today's date."""
        today = datetime.now().strftime("%m-%d")  # Format: MM-DD
        # Filter birthdays where the MM-DD part matches today
        today_birthdays = [name for name, dob in self.birthdays.items() if dob and dob.endswith(today)]
        return today_birthdays if today_birthdays else []

    def days_and_months_until_birthday(self, birthday):
        """
        Calculates the months and days remaining until a given birthday.
        Assumes birthday format is MM-DD.
        :param birthday: Birthday string in "MM-DD" format.
        :return: Tuple of (months_left, exact_days_left).
        """
        try:
            today = datetime.now()
            # Extract month and day from the birthday string
            bday_month, bday_day = map(int, birthday.split('-'))

            # Create a datetime object for the birthday in the current year
            bday_date_this_year = datetime(today.year, bday_month, bday_day)

            # If the birthday has already passed this year, set it for next year
            if bday_date_this_year < today:
                bday_date = bday_date_this_year.replace(year=today.year + 1)
            else:
                bday_date = bday_date_this_year

            delta = bday_date - today  # Calculate the time difference
            days_left = delta.days  # Total days remaining

            # Approximate months and remaining days
            months_left = days_left // 30
            exact_days_left = days_left % 30

            return months_left, exact_days_left
        except ValueError:
            # Handle cases where birthday string is not in expected format
            return 0, 0

    def calculate_progress(self):
        """
        Calculates study progress based on the 'Marks' sheet in the schedule Excel file.
        :return: Tuple of (total_marks, percentage_out_of_40, needed_to_reach_27).
        """
        if self.df_marks is None or self.df_marks.empty or not all(
                col in self.df_marks.columns for col in ["Test-01", "Test-02", "Test-03", "Assignment"]):
            return 0, 0, 0  # Return zeros if DataFrame is not ready or columns are missing

        # Sum all marks from the specified columns
        total_marks = self.df_marks[["Test-01", "Test-02", "Test-03", "Assignment"]].sum().sum()

        # Calculate max possible marks (assuming 6 subjects * 4 categories * 50 marks/category = 1200)
        # Adjust 50 if your tests/assignments are out of a different total
        max_marks = len(self.df_marks) * 4 * 50  # Assuming 50 marks per test/assignment

        # Avoid division by zero
        percentage = (total_marks / max_marks) * 40 if max_marks > 0 else 0

        # Marks needed to reach 27 out of 40 (which is (27/40) * max_marks)
        target_marks_27_of_40 = (27 / 40) * max_marks
        needed = target_marks_27_of_40 - total_marks

        return total_marks, percentage, max(0, needed)  # Ensure 'needed' is not negative

    def update_marks(self, test_name, subject_scores):
        """
        Updates specific test or assignment marks in the 'Marks' sheet of the schedule Excel.
        :param test_name: Name of the test/assignment column (e.g., "Test-01").
        :param subject_scores: Dictionary of subject names and their scores.
        :return: A string summarizing the update and current progress.
        """
        response = []
        try:
            # Ensure the test_name column exists, if not, add it
            if test_name not in self.df_marks.columns:
                self.df_marks[test_name] = 0  # Add new column with default 0s
                response.append(f"Added new column for '{test_name}'.")

            for subject, score in subject_scores.items():
                # Find the row index for the given subject
                subject_idx = self.df_marks[self.df_marks["Subject"] == subject].index
                if not subject_idx.empty:
                    # Update the score at the found index and column
                    self.df_marks.loc[subject_idx[0], test_name] = score
                else:
                    response.append(
                        f"Warning: Subject '{subject}' not found in marks sheet. Skipping update for this subject.")

            # Save the updated DataFrame back to the 'Marks' sheet
            # Use mode='a' (append) and if_sheet_exists='replace' to update existing sheet
            with pd.ExcelWriter(self.schedule_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.df_marks.to_excel(writer, sheet_name='Marks', index=False)

            response.append(f"Okay, I’ve updated the marks for {test_name}.")
            response.append("Here’s the new info:\n" + self.df_marks.to_string())  # Display current marks table

            total, percent, needed = self.calculate_progress()  # Recalculate progress
            response.append(
                f"You’ve got {total} out of 1200 now, which is {percent:.1f} out of 40. You need {needed} more to reach 27 out of 40.")
        except Exception as e:
            response.append(f"Error updating marks: {e}")
        return "\n".join(response)

    def mark_studyschedule(self, user_input):
        """
        Handles commands related to study schedules and marks.
        :param user_input: The user's input string.
        :return: A string response.
        """
        response = []
        today = datetime.now()
        date_str = today.strftime("%b %d").replace(" 0", " ")  # e.g., "Jul 19"

        if "today" in user_input.lower() and "schedule" in user_input.lower():
            if date_str in self.schedule:
                response.append(f"Here’s what you’ve got today, {date_str}:")
                response.extend(self.schedule[date_str].split('; '))  # Split events by semicolon
            else:
                response.append(
                    f"Looks like there’s nothing planned for {date_str}. Want me to help you set something up?")
        elif "schedule for" in user_input.lower():
            try:
                date_part_match = re.search(r"schedule for (.*)", user_input.lower())
                if date_part_match:
                    date_part = date_part_match.group(1).strip()
                    # Attempt to parse various date formats (e.g., "Jul 19", "July 19")
                    try:
                        target_date = datetime.strptime(date_part, "%b %d")
                    except ValueError:
                        target_date = datetime.strptime(date_part, "%B %d")

                    date_str_formatted = target_date.strftime("%b %d").replace(" 0", " ")
                    if date_str_formatted in self.schedule:
                        response.append(f"Here’s the plan for {date_str_formatted}:")
                        response.extend(self.schedule[date_str_formatted].split('; '))
                    else:
                        response.append(f"No plans for {date_str_formatted} yet. Should I add something for you?")
                else:
                    response.append("Hmm, that date doesn’t look right. Try saying 'schedule for Feb 26'.")
            except ValueError:
                response.append("Hmm, that date doesn’t look right. Try saying 'schedule for Feb 26'.")
        elif "schedule" in user_input.lower():  # Default to tomorrow's schedule
            target_date = today + timedelta(days=1)
            date_str_formatted = target_date.strftime("%b %d").replace(" 0", " ")
            if date_str_formatted in self.schedule:
                response.append(f"Here’s what’s coming up tomorrow, {date_str_formatted}:")
                response.extend(self.schedule[date_str_formatted].split('; '))
            else:
                response.append(f"There’s nothing set for {date_str_formatted} yet. Want me to help plan it?")
        elif "mark" in user_input.lower():
            response.append("Here’s how your grades are looking right now:")
            response.append(self.df_marks.to_string())  # Display the marks DataFrame
            total, percent, needed = self.calculate_progress()
            response.append(
                f"You’re at {total} out of 1200, which is {percent:.1f} out of 40. You need {needed} more to get to 27 out of 40.")
            if "update" in user_input.lower():
                marks_input = simpledialog.askstring("Input",
                                                     "Tell me the marks like 'Test-02: Subject 1=40, Subject 2=45'",
                                                     parent=self.master)
                if marks_input:
                    try:
                        # Regex to parse "Test-XX:" or "Assignment:"
                        test_name_match = re.match(r"(Test-\d+|Assignment):\s*(.*)", marks_input, re.IGNORECASE)
                        if test_name_match:
                            test_name = test_name_match.group(1).strip()
                            scores_part = test_name_match.group(2).strip()
                            subject_scores = {}
                            for item in scores_part.split(","):
                                if "=" in item:
                                    subject, score = item.split("=")
                                    subject_scores[subject.strip()] = int(score.strip())
                                else:
                                    raise ValueError(f"Invalid score format in '{item}'. Expected 'Subject=Score'.")
                            response.append(self.update_marks(test_name, subject_scores))
                        else:
                            raise ValueError(
                                "Invalid input format. Expected 'Test-XX: Subject=Score,...' or 'Assignment: Subject=Score,...'")
                    except ValueError as ve:
                        response.append(
                            f"Oops, something’s wrong with that format: {ve}. Try 'Test-02: Subject 1=40, Subject 2=45'.")
                    except Exception as e:
                        response.append(f"An unexpected error occurred while parsing marks: {e}")
                else:
                    response.append(
                        "I need some marks to update! Try again. Format: 'Test-02: Subject 1=40, Subject 2=45'")
            else:
                response.append("If you want to change anything, just say 'update mark'.")
        return "\n".join(response)

    def teach_english(self, level=None):
        """
        Provides English language practice sentences from the loaded English data.
        :param level: The specific level to teach (integer), or None for a random level.
        :return: A string containing English sentences and explanations.
        """
        if not self.english_levels:
            return f"{self.ai_name}: I don't have English teaching data loaded. Please check the '{self.ai_name.lower()}_english_data.xlsx' file."

        response_lines = []
        selected_level_data = []
        chosen_level_name = ""

        if level is None:
            # Pick a random level if none specified
            response_lines.append(f"{self.ai_name}: I’ll pick a random level for you today. Here are five sentences:")
            chosen_level_name = random.choice(list(self.english_levels.keys()))
            selected_level_data = self.english_levels[chosen_level_name]
        else:
            chosen_level_name = f"Level {level}"
            if chosen_level_name in self.english_levels:
                response_lines.append(
                    f"{self.ai_name}: Here’s your English practice for {chosen_level_name}—five sentences:")
                selected_level_data = self.english_levels[chosen_level_name]
            else:
                # List available levels if the requested level is not found
                available_levels = sorted([s.replace('Level ', '') for s in self.english_levels.keys()])
                return f"{self.ai_name}: Hey, level '{level}' not found. I have levels: {', '.join(available_levels)}. Try something like 'teach me English level 3'!"

        # Provide up to 5 random sentences from the selected level
        if len(selected_level_data) >= 5:
            for sentence, explanation in random.sample(selected_level_data, 5):
                response_lines.append(f"- {sentence} ({explanation})")
        else:
            # If less than 5 sentences, provide all available
            for sentence, explanation in selected_level_data:
                response_lines.append(f"- {sentence} ({explanation})")
            response_lines.append(f"Note: This level only has {len(selected_level_data)} sentences available.")

        response_lines.append(f"{self.ai_name}: Try these out, and you’ll get the hang of it!")
        return "\n".join(response_lines)

    def generate_response(self, message):
        """
        Generates an AI response based on the user's input message.
        This is the core logic for the chatbot.
        :param message: The user's input string.
        :return: The AI's response string.
        """
        # --- Personal Information Extraction and Storage ---
        name_match = re.search(r"(?:my name is|I'm)\s+([A-Za-z\s]+)", message, re.IGNORECASE)
        birthday_match = re.search(r"(?:my birthday is|born on)\s+([A-Za-z0-9\s,/-]+)", message, re.IGNORECASE)
        address_match = re.search(r"(?:I live in|my address is)\s+(.+)", message, re.IGNORECASE)
        number_match = re.search(r"(?:my number is|phone number is)\s+([0-9\s\-\+]+)", message, re.IGNORECASE)

        # --- Information Retrieval Queries ---
        who_match = re.search(r"who\s*(am\s*I|is\s*this)\s*\??", message, re.IGNORECASE)

        # --- General Commands ---
        joke_match = re.search(r"tell\s*(me\s*a\s*joke|joke)", message, re.IGNORECASE)

        # Check for previous response to avoid repetition for exact same query
        previous_response = self.find_previous_response(message)
        if previous_response and previous_response.startswith(self.ai_name):  # Ensure it's an AI's response
            return previous_response

        # --- Process Commands in Order of Priority ---
        if name_match:
            self.user_info["name"] = name_match.group(1).strip()
            return f"{self.ai_name}: Nice to meet you, {self.user_info['name']}! What's your birthday?"
        elif birthday_match:
            try:
                bday_raw = birthday_match.group(1).strip()
                # Attempt to normalize common date formats (e.g., "03-04", "March 4", "03/04/1990")
                if re.match(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", bday_raw):  # e.g., 03-04-1990 or 3/4/90
                    parts = re.split(r"[-/]", bday_raw)
                    if len(parts) == 3:  # MM-DD-YYYY or DD-MM-YYYY
                        try:
                            dt_obj = datetime.strptime(bday_raw, "%m-%d-%Y")
                        except ValueError:
                            dt_obj = datetime.strptime(bday_raw, "%d-%m-%Y")  # Try DD-MM-YYYY
                        bday_formatted = dt_obj.strftime("%m-%d")  # Store as MM-DD
                    elif len(parts) == 2:  # MM-DD
                        bday_formatted = bday_raw
                    else:
                        raise ValueError("Unsupported date format.")
                elif re.match(r"[A-Za-z]+\s+\d{1,2}", bday_raw):  # e.g., March 4
                    dt_obj = datetime.strptime(bday_raw, "%B %d")
                    bday_formatted = dt_obj.strftime("%m-%d")
                else:
                    raise ValueError("Unsupported date format.")

                self.user_info["birthday"] = bday_formatted
                if self.user_info["name"]:  # Only save if user's name is known
                    self.birthdays[self.user_info["name"]] = bday_formatted
                    self.save_birthdays()
                return f"{self.ai_name}: I'll remember that! Where do you live?"
            except ValueError:
                return f"{self.ai_name}: Hmm, I couldn't understand that birthday format. Please use MM-DD, e.g., '03-04' or 'March 4'."
        elif address_match:
            self.user_info["address"] = address_match.group(1).strip()
            return f"{self.ai_name}: Sounds like a great place! What's your phone number?"
        elif number_match:
            self.user_info["number"] = number_match.group(1).strip()
            return f"{self.ai_name}: Got it! Anything else you'd like to share?"
        elif who_match:
            if self.user_info["name"]:
                return f"{self.ai_name}: You're {self.user_info['name']}! Anything else you'd like to share?"
            else:
                return f"{self.ai_name}: I don't know your name yet! Please tell me, what's your name?"
        elif joke_match:
            joke = random.choice(self.joke_list)
            return joke
        elif "what is my name" in message.lower():
            return f"{self.ai_name}: Your name’s {self.user_info.get('name', 'not set yet')}, right?"
        elif "what is my birthday" in message.lower():
            if self.user_info["name"] and self.user_info["name"] in self.birthdays:
                months_left, exact_days_left = self.days_and_months_until_birthday(
                    self.birthdays[self.user_info["name"]])
                if months_left > 0:
                    return f"{self.ai_name}: Hey {self.user_info['name']}, your birthday’s coming up in {months_left} months and {exact_days_left} days!"
                else:
                    return f"{self.ai_name}: {self.user_info['name']}, it’s just {exact_days_left} days until your birthday—awesome!"
            else:
                return f"{self.ai_name}: I don’t have your birthday yet. Say 'store my birthday [name],[MM-DD]' to add it!"
        elif "what is my address" in message.lower():
            return f"{self.ai_name}: Your address is {self.user_info.get('address', 'not set yet')}, as far as I know."
        elif "what is my number" in message.lower() or "what is my phone number" in message.lower():
            return f"{self.ai_name}: Your phone number is {self.user_info.get('number', 'not set yet')}—that’s what I’ve got!"
        elif "who's birthday today" in message.lower():
            today_birthdays = self.check_birthdays()
            return f"{self.ai_name}: Today’s birthdays are {', '.join(today_birthdays)}—time to celebrate!" if today_birthdays else f"{self.ai_name}: No birthdays today, looks like a chill day."
        elif "schedule" in message.lower() or "mark" in message.lower():
            return self.mark_studyschedule(message)
        elif "teach me english" in message.lower() or "english lesson" in message.lower():
            level_match = re.search(r"level\s+(\d+)", message, re.IGNORECASE)
            if level_match:
                try:
                    level = int(level_match.group(1))
                    return self.teach_english(level)
                except ValueError:
                    return f"{self.ai_name}: I need a number for the level—like 'teach me English level 5'. Give it a try!"
            else:
                return self.teach_english()
        elif "open youtube" in message.lower():
            self.open_youtube()
            return f"{self.ai_name}: YouTube’s up—enjoy some videos!"
        elif "open google" in message.lower():
            webbrowser.open("https://www.google.com")
            return f"{self.ai_name}: Google’s ready—go search whatever you want!"
        elif "music" in message.lower() or "song" in message.lower():
            self.play_music()
            return f"{self.ai_name}: Playing some tunes for you now!"
        elif "time" in message.lower():
            strTime = datetime.now().strftime("%H:%M:%S")
            return f"{self.ai_name}: Hey, it’s {strTime} right now—time flies, doesn’t it?"
        elif "open notepad" in message.lower():
            self.open_notepad()
            return f"{self.ai_name}: Notepad’s open—start writing whatever you like!"
        elif "play games" in message.lower():
            # IMPORTANT: Replace "path_to_your_game" with the actual path to your game executable
            # Example: os.startfile("C:\\Games\\MyGame\\MyGame.exe")
            # os.startfile("path_to_your_game")
            return f"{self.ai_name}: Launching a game for you—have fun! (Note: Game path needs to be configured in the code)"
        elif "search" in message.lower():
            search_query = message.replace("search", "").strip()
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            return f"{self.ai_name}: Searching Google for '{search_query}'—it’ll pop up soon!"
        elif "wikipedia" in message.lower():
            try:
                search_term = message.replace("wikipedia", "").strip()
                self.display_async(f"{self.ai_name}: Searching Wikipedia for '{search_term}'...", "ai")
                results = wikipedia.summary(search_term, sentences=2)
                return f"{self.ai_name}: Here’s what I found: {results}"
            except wikipedia.exceptions.DisambiguationError as e:
                return f"{self.ai_name}: I found multiple results for '{search_term}'. Can you be more specific? Options include: {', '.join(e.options[:5])}..."
            except wikipedia.exceptions.PageError:
                return f"{self.ai_name}: Sorry, I couldn’t find any Wikipedia page for '{search_term}'. Can you rephrase?"
            except Exception as e:
                return f"{self.ai_name}: Sorry, I couldn’t search Wikipedia—something went wrong: {str(e)}."
        elif message.lower().startswith("open ") or message.lower().startswith("play "):
            return self.process_open_command(message)
        elif "minimize window" in message.lower():
            return self.window_control("minimize")
        elif "maximize window" in message.lower():
            return self.window_control("maximize")
        elif "restore window" in message.lower() or "restore down" in message.lower():
            return self.window_control("restore")
        elif "close window" in message.lower():
            return self.window_control("close")
        elif "volume up" in message.lower():
            return self.volume_up()
        elif "volume down" in message.lower():
            return self.volume_down()
        elif "mute volume" in message.lower():
            return self.mute_volume()
        elif "play video" in message.lower() or "pause video" in message.lower():
            return self.play_pause()
        elif "next track" in message.lower():
            return self.next_track()
        elif "previous track" in message.lower():
            return self.previous_track()
        elif "brightness up" in message.lower():
            return self.brightness_up()
        elif "brightness down" in message.lower():
            return self.brightness_down()
        elif "cmd" in message.lower() or "prompt" in message.lower():
            self.cmd()
            return f"{self.ai_name}: There you go—the command prompt’s ready!"
        elif "open spotify" in message.lower():
            self.open_spotify()
            return f"{self.ai_name}: Spotify’s ready—pick your tunes!"
        elif "open whatsapp" in message.lower():
            self.open_whatsapp()
            return f"{self.ai_name}: WhatsApp’s up—start chatting!"
        elif "exit" in message.lower() or "bye" in message.lower() or "stop" in message.lower():
            self.speak_async(f"{self.ai_name}: See you later—take care!")
            # Correct way to close the root window from a child frame
            self.master.winfo_toplevel().destroy()
            return f"{self.ai_name}: Shutting down..."
        elif "hello" in message.lower() or "hi" in message.lower() or "hey" in message.lower():
            return random.choice([
                f"{self.ai_name}: Hi! How’s your day going so far?",
                f"{self.ai_name}: Hey there! What’s up with you today?",
                f"{self.ai_name}: Hello! Nice to chat—what’s on your mind?"
            ])
        elif "thank you" in message.lower() or "thanks" in message.lower():
            return random.choice([
                f"{self.ai_name}: You’re welcome! Happy to help anytime.",
                f"{self.ai_name}: No problem! What else can I do for you?",
                f"{self.ai_name}: Anytime! What’s next on your list?"
            ])
        elif "how are you" in message.lower():
            return f"{self.ai_name}: I’m doing great, thanks! How about you?"
        elif "who created you" in message.lower() or "made you" in message.lower():
            return f"{self.ai_name}: I was made by someone smart to help folks like you. Pretty neat, huh?"
        elif "what can you do" in message.lower():
            return random.choice([
                f"{self.ai_name}: I can help with time, schedules, music, and storing your info!",
                f"{self.ai_name}: I can open stuff, check birthdays, or even tell a joke—whatever you need!",
                f"{self.ai_name}: I’m here for chats, reminders, or quick tasks—what’s up?"
            ]) + "\nTry saying 'help', 'store my name', or 'what’s the time'—I’ve got you covered!"
        elif "motivate" in message.lower():
            return random.choice([
                f"{self.ai_name}: You’ve got this! Keep going—you’re doing awesome!",
                f"{self.ai_name}: Every step you take is a win—don’t stop now!",
                f"{self.ai_name}: Hey, you’re amazing—keep pushing and you’ll get there!"
            ])
        elif "help" in message.lower():
            return f"{self.ai_name}: I can store your personal stuff! Here’s how:\n" \
                   "- 'store my name [your name]'\n" \
                   "- 'store my age [your age]'\n" \
                   "- 'store my birthday [name],[MM-DD]'\n" \
                   "- 'store my favorite dish [dish]'\n" \
                   "- 'store my friend name [friend]'\n" \
                   "- 'store my email [email]'\n" \
                   "- 'store my phone [number]'\n" \
                   "- 'store my address [address]'\n" \
                   "Ask 'what is my [thing]' to check it later!"

        return f"{self.ai_name}: I see. Tell me more."  # Default fallback response

    def handle_send(self):
        """Processes the user's text input from the entry field."""
        message = self.entry.get().strip()  # Get and strip whitespace from message
        if not message:
            return  # Do nothing if message is empty

        self.display_async(f"You: {message}", "user")  # Display user message
        response = self.generate_response(message)  # Get AI's response
        self.display_async(response, "ai")  # Display AI's response
        self.speak_async(response)  # Speak the response

        # Store conversation to Excel
        self.store_to_excel(
            "User",  # Use a generic "User" string
            message,
            response.replace(f"{self.ai_name}: ", ""),  # Store response without AI name prefix
            name=self.user_info["name"],
            birthday=self.user_info["birthday"],
            address=self.user_info["address"],
            number=self.user_info["number"],
            **{f"custom{i + 1}": "" for i in range(self.custom_column_count)}  # Pass empty for custom for now
        )
        self.entry.delete(0, tk.END)  # Clear the input field

    def listen_continuous(self):
        """
        FIXED: Continuously listens for voice commands in a background thread.
        The `with sr.Microphone() as source:` block now correctly wraps the `while` loop
        to keep the microphone stream open.
        """
        self.listening = True
        self.display_async(f"{self.ai_name}: Calibrating microphone, please wait...", "ai")
        self.speak_async("Calibrating microphone, please wait.")

        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                self.display_async(f"{self.ai_name}: Listening...", "ai")
                self.speak_async("I'm listening.")

                while self.listening:
                    try:
                        # Listen for audio. The timeout prevents it from blocking indefinitely.
                        audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                        # Recognize speech using Google Web Speech API
                        user_input = self.recognizer.recognize_google(audio)

                        # Process the recognized input in the main thread
                        self.master.after(0, self.process_voice_command, user_input)

                    except sr.WaitTimeoutError:
                        # This is normal. It means no speech was detected in the timeout period.
                        # The loop will just continue to the next iteration.
                        pass
                    except sr.UnknownValueError:
                        # This means speech was detected but couldn't be understood.
                        if self.listening:
                            self.display_async(f"{self.ai_name}: Sorry, I didn’t catch that.", "ai")
                    except sr.RequestError as e:
                        # This means there was an issue with the API request (e.g., no internet).
                        if self.listening:
                            error_msg = f"API Error: Could not request results; {e}"
                            self.display_async(f"{self.ai_name}: {error_msg}", "ai")
                            self.speak_async("There seems to be an issue with the connection.")
                            self.listening = False  # Stop listening on API error
                    except Exception as e:
                        if self.listening:
                            self.display_async(f"{self.ai_name}: An unexpected error occurred: {e}", "ai")
                            self.listening = False  # Stop on other errors

        except (OSError, AttributeError):
            self.display_async(
                f"{self.ai_name}: Microphone not found. Please ensure it's connected and drivers are installed.", "ai")
            self.speak_async("Microphone not found.")
            self.listening = False

    def process_voice_command(self, user_input):
        """
        This function is called from the main thread to process the recognized text,
        ensuring GUI updates are safe.
        """
        self.display_async(f"You (Voice): {user_input}", "user")
        response = self.generate_response(user_input)
        self.display_async(response, "ai")
        self.speak_async(response)
        # Store conversation to Excel
        self.store_to_excel(
            "User (Voice)",
            user_input,
            response.replace(f"{self.ai_name}: ", ""),
            name=self.user_info["name"],
            birthday=self.user_info["birthday"],
            address=self.user_info["address"],
            number=self.user_info["number"]
        )

    def start_listening(self):
        """Starts the continuous voice listening thread."""
        if not self.listening:
            # Check for microphones before starting the thread
            if not sr.Microphone.list_microphone_names():
                messagebox.showerror("Microphone Error",
                                     "No microphone found. Please connect a microphone to use voice commands.")
                return

            self.listen_thread = threading.Thread(target=self.listen_continuous, daemon=True)
            self.listen_thread.start()
        else:
            self.display_async(f"{self.ai_name}: I'm already listening!", "ai")

    def stop_listening(self):
        """Stops the continuous voice listening thread."""
        if self.listening:
            self.listening = False
            self.display_async(f"{self.ai_name}: Voice input disabled.", "ai")
            self.speak_async("Voice input disabled.")
        else:
            self.display_async(f"{self.ai_name}: I'm not currently listening.", "ai")

    # --- NEW, IMPROVED FILE SEARCH AND OPEN LOGIC (NON-BLOCKING) ---
    def normalize_text(self, text):
        """Converts text to a standard format for easy matching."""
        return text.lower().replace(" ", "").replace("-", "").replace("_", "")

    def search_local_files(self, query):
        """
        Searches the specified drives for files matching the query.
        This function is designed to be run in a background thread.
        """
        # --- CONFIGURATION ---
        SEARCH_DRIVES = ['C:/', 'E:/']
        FILE_TYPES = [
            ".mp3", ".wav", ".flac", ".m4a", ".mp4", ".mkv", ".avi", ".mov",
            ".pdf", ".docx", ".txt", ".jpg", ".png", ".jpeg", ".pptx", ".xlsx"
        ]
        # --- END CONFIGURATION ---

        matched_files = []
        normalized_query = self.normalize_text(query)

        for drive in SEARCH_DRIVES:
            # Use self.master.after to safely update the GUI from this thread
            self.master.after(0, self.display_async, f"Scanning drive {drive}...", "ai")

            if not os.path.exists(drive):
                print(f"Warning: Drive {drive} not found. Skipping.")
                continue

            for root, _, files in os.walk(drive, topdown=True):
                for file in files:
                    if any(file.lower().endswith(ext) for ext in FILE_TYPES):
                        base_filename = os.path.splitext(file)[0]
                        normalized_filename = self.normalize_text(base_filename)

                        if normalized_query in normalized_filename:
                            full_path = os.path.join(root, file)
                            match_data = (file, len(base_filename), full_path)
                            matched_files.append(match_data)
        return matched_files

    def find_best_match(self, matches):
        """Finds the best match (shortest filename) from a list of files."""
        if not matches:
            return None
        matches.sort(key=lambda x: x[1])
        return matches[0]

    def open_file_crossplatform(self, path):
        """Opens a file using the system's default application (cross-platform)."""
        try:
            os.startfile(path)
        except AttributeError:
            try:
                if sys.platform == "darwin":  # macOS
                    subprocess.run(["open", path], check=True)
                else:  # Linux
                    subprocess.run(["xdg-open", path], check=True)
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self.display_async(f"Error: Could not open file. Your OS may not have a default handler. {e}", "ai")
                return False
        except Exception as e:
            self.display_async(f"An error occurred while opening the file: {e}", "ai")
            return False
        return True

    def process_open_command(self, input_text):
        """
        Handles commands to open/play files by starting a background search thread.
        This keeps the GUI responsive.
        """
        match = re.match(r"(?:open|play)\s+(.*)", input_text, re.IGNORECASE)
        if not match:
            return f"{self.ai_name}: Oops, you need to say 'open [file name]' or 'play [song name]' so I know what to do!"

        query = match.group(1).strip()
        if not query:
            return f"{self.ai_name}: What would you like me to open or play?"

        # Start the search in a separate thread to avoid freezing the GUI
        search_thread = threading.Thread(target=self._search_and_open_worker, args=(query,), daemon=True)
        search_thread.start()

        # Immediately return a message to the user so the GUI doesn't wait
        return f"{self.ai_name}: On it! Searching for '{query}' in the background. I'll let you know what I find."

    def _search_and_open_worker(self, query):
        """
        This function runs in a background thread to search for files.
        It calls back to the main GUI thread when done.
        """
        found_files = self.search_local_files(query)

        # Schedule the result handler to run on the main GUI thread
        self.master.after(0, self._handle_search_result, query, found_files)

    def _handle_search_result(self, query, found_files):
        """
        This function runs on the main GUI thread to process search results
        and update the UI safely.
        """
        if not found_files:
            response = f"\n{self.ai_name}: Sorry, I couldn’t find anything matching '{query}' on your local drives."
            self.display_async(response, "ai")
            self.speak_async(response)
            return

        best_match = self.find_best_match(found_files)
        best_match_name = best_match[0]
        best_match_path = best_match[2]

        response1 = f"\n{self.ai_name}: Found {len(found_files)} match(es). The best one is '{best_match_name}'."
        self.display_async(response1, "ai")
        self.speak_async(response1)

        if self.open_file_crossplatform(best_match_path):
            response2 = f"{self.ai_name}: Opening it for you now!"
        else:
            response2 = f"{self.ai_name}: I found the file but couldn't open it. There might be an issue with your system's file associations."

        self.display_async(response2, "ai")
        self.speak_async(response2)

    # --- END NEW FILE SEARCH LOGIC ---

    def window_control(self, action):
        """
        Performs basic window control actions (minimize, maximize, restore, close)
        on the currently active window (Windows-specific).
        :param action: The action to perform.
        :return: A string response.
        """
        try:
            user32 = ctypes.windll.user32
            hwnd = user32.GetForegroundWindow()  # Get handle of the active window

            if action == "minimize":
                user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
                return f"{self.ai_name}: I’ve minimized the window for you."
            elif action == "maximize":
                user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                return f"{self.ai_name}: The window’s maximized now."
            elif action == "restore":
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                return f"{self.ai_name}: I’ve restored the window to normal size."
            elif action == "close":
                user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE message
                return f"{self.ai_name}: Window’s closed now."
            else:
                return f"{self.ai_name}: I don't recognize that window action."
        except Exception as e:
            return f"{self.ai_name}: I encountered an error trying to control the window: {e}"

    def volume_up(self):
        """Increases the system volume (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xAF, 0, 0, 0)  # VK_VOLUME_UP
            ctypes.windll.user32.keybd_event(0xAF, 0, 2, 0)  # KEYEVENTF_KEYUP
            return f"{self.ai_name}: I’ve turned the volume up a bit for you."
        except Exception as e:
            return f"{self.ai_name}: Couldn't adjust volume up: {e}"

    def volume_down(self):
        """Decreases the system volume (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xAE, 0, 0, 0)  # VK_VOLUME_DOWN
            ctypes.windll.user32.keybd_event(0xAE, 0, 2, 0)
            return f"{self.ai_name}: I’ve lowered the volume for you."
        except Exception as e:
            return f"{self.ai_name}: Couldn't adjust volume down: {e}"

    def mute_volume(self):
        """Mutes/unmutes the system volume (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xAD, 0, 0, 0)  # VK_VOLUME_MUTE
            ctypes.windll.user32.keybd_event(0xAD, 0, 2, 0)
            return f"{self.ai_name}: The sound’s muted now—nice and quiet."
        except Exception as e:
            return f"{self.ai_name}: Couldn't mute volume: {e}"

    def play_pause(self):
        """Toggles media play/pause (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xB3, 0, 0, 0)  # VK_MEDIA_PLAY_PAUSE
            ctypes.windll.user32.keybd_event(0xB3, 0, 2, 0)
            return f"{self.ai_name}: I’ve toggled play or pause for you—your choice!"
        except Exception as e:
            return f"{self.ai_name}: Couldn't toggle play/pause: {e}"

    def next_track(self):
        """Skips to the next media track (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xB0, 0, 0, 0)  # VK_MEDIA_NEXT_TRACK
            ctypes.windll.user32.keybd_event(0xB0, 0, 2, 0)
            return f"{self.ai_name}: Moving on to the next track now."
        except Exception as e:
            return f"{self.ai_name}: Couldn't skip to next track: {e}"

    def previous_track(self):
        """Goes to the previous media track (Windows-specific)."""
        try:
            ctypes.windll.user32.keybd_event(0xB1, 0, 0, 0)  # VK_MEDIA_PREV_TRACK
            ctypes.windll.user32.keybd_event(0xB1, 0, 2, 0)
            return f"{self.ai_name}: Going back to the last track for you."
        except Exception as e:
            return f"{self.ai_name}: Couldn't go to previous track: {e}"

    def brightness_up(self):
        """
        Increases screen brightness (Windows-specific, uses PowerShell).
        May require administrative privileges or specific display drivers.
        """
        try:
            # Uses WMI to set brightness. Value 100 is max.
            subprocess.run(["powershell",
                            "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,100)"],
                           check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            return f"{self.ai_name}: I’ve brightened the screen for you."
        except subprocess.CalledProcessError as e:
            return f"{self.ai_name}: Couldn't adjust brightness. This feature might require specific permissions or drivers. Error: {e}"
        except FileNotFoundError:
            return f"{self.ai_name}: PowerShell not found. Cannot adjust brightness."
        except Exception as e:
            return f"{self.ai_name}: An unexpected error occurred while adjusting brightness: {e}"

    def brightness_down(self):
        """
        Decreases screen brightness (Windows-specific, uses PowerShell).
        May require administrative privileges or specific display drivers.
        """
        try:
            # Uses WMI to set brightness. Value 10 is a low brightness level.
            subprocess.run(["powershell",
                            "(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,10)"],
                           check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            return f"{self.ai_name}: The screen’s dimmed a bit now."
        except subprocess.CalledProcessError as e:
            return f"{self.ai_name}: Couldn't adjust brightness. This feature might require specific permissions or drivers. Error: {e}"
        except FileNotFoundError:
            return f"{self.ai_name}: PowerShell not found. Cannot adjust brightness."
        except Exception as e:
            return f"{self.ai_name}: An unexpected error occurred while adjusting brightness: {e}"

    def play_music(self):
        """
        Plays a random music file from a specified directory.
        IMPORTANT: Update `music_dir` to your actual music folder path.
        """
        # <<< IMPORTANT: UPDATE THIS PATH TO YOUR MUSIC FOLDER >>>
        music_dir = "E:\\Music"
        if not os.path.exists(music_dir):
            self.display_async(
                f"{self.ai_name}: Music folder '{music_dir}' not found. Please update the path in the code.", "ai")
            self.speak_async(f"{self.ai_name}: Music folder not found.")
            return

        # Filter for common audio file extensions
        audio_files = [f for f in os.listdir(music_dir) if
                       f.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.m4a'))]

        if audio_files:
            song = random.choice(audio_files)
            try:
                self.open_file_crossplatform(os.path.join(music_dir, song))
                self.display_async(f"{self.ai_name}: Playing {song} for you—enjoy!", "ai")
            except Exception as e:
                self.display_async(f"{self.ai_name}: Couldn't play {song}. Error: {e}", "ai")
        else:
            self.display_async(f"{self.ai_name}: No audio files found in '{music_dir}'.", "ai")

    def cmd(self):
        """Opens the Command Prompt (Windows-specific)."""
        self.speak_async(f"{self.ai_name}: Opening the command prompt for you!")
        try:
            subprocess.Popen("cmd.exe")  # Use Popen for non-blocking execution
        except FileNotFoundError:
            self.display_async(f"{self.ai_name}: Command Prompt application not found.", "ai")
        except Exception as e:
            self.display_async(f"{self.ai_name}: Error opening Command Prompt: {e}", "ai")

    def open_notepad(self):
        """Opens the Notepad application (Windows-specific)."""
        self.display_async(f"{self.ai_name}: Notepad’s up—go ahead and write!", "ai")
        try:
            os.startfile("notepad.exe")
        except FileNotFoundError:
            self.display_async(f"{self.ai_name}: Notepad application not found.", "ai")
        except Exception as e:
            self.display_async(f"{self.ai_name}: Error opening Notepad: {e}", "ai")

    def open_youtube(self):
        """Opens YouTube in the default web browser."""
        self.display_async(f"{self.ai_name}: YouTube’s loading up now!", "ai")
        webbrowser.open("https://www.youtube.com")

    def google_search(self):
        """Performs a Google search using the text from the input field."""
        search_query = self.entry.get().strip()
        if not search_query:
            self.display_async(f"{self.ai_name}: What do you want me to search on Google?", "ai")
            self.speak_async(f"{self.ai_name}: What do you want me to search on Google?")
            return
        self.display_async(f"{self.ai_name}: Searching Google for '{search_query}'!", "ai")
        webbrowser.open(f"https://www.google.com/search?q={search_query}")

    def take_screenshot(self):
        """
        Captures a screenshot and saves it to the user's Pictures/Screenshots folder.
        Requires `pyautogui` library.
        """
        try:
            file_name = simpledialog.askstring("Input", "What do you want to name this screenshot?", parent=self.master)
            if file_name:
                # Save to a common screenshots directory
                screenshot_dir = os.path.join(os.path.expanduser("~"), "Pictures", "Screenshots")
                os.makedirs(screenshot_dir, exist_ok=True)  # Ensure directory exists
                full_path = os.path.join(screenshot_dir, f"{file_name}.png")
                pyautogui.screenshot(full_path)  # Take and save screenshot
                self.display_async(f"{self.ai_name}: Saved your screenshot as {full_path}!", "ai")
            else:
                self.display_async(f"{self.ai_name}: No name provided, screenshot not saved.", "ai")
        except Exception as e:
            self.display_async(f"{self.ai_name}: Failed to take screenshot: {e}", "ai")
            messagebox.showerror("Screenshot Error", f"Failed to take screenshot: {e}", parent=self.master)

    def open_spotify(self):
        """
        Opens the Spotify desktop application.
        IMPORTANT: Update `spotify_path` to your actual Spotify executable path.
        """
        # <<< IMPORTANT: UPDATE THIS PATH TO YOUR SPOTIFY EXECUTABLE >>>
        spotify_path = "C:\\Users\\DELL\\AppData\\Roaming\\Spotify\\Spotify.exe"
        try:
            subprocess.Popen([spotify_path])
            self.display_async(f"{self.ai_name}: Spotify’s ready—pick your tunes!", "ai")
        except FileNotFoundError:
            self.display_async(f"{self.ai_name}: Spotify not found at '{spotify_path}'—check the path!", "ai")
        except Exception as e:
            self.display_async(f"{self.ai_name}: Spotify issue: {e}.", "ai")

    def open_whatsapp(self):
        """
        Opens the WhatsApp desktop application.
        IMPORTANT: Update `whatsapp_path` to your actual WhatsApp executable path.
        """
        # <<< IMPORTANT: UPDATE THIS PATH TO YOUR WHATSAPP DESKTOP EXECUTABLE >>>
        whatsapp_path = "C:\\Users\\DELL\\AppData\\Local\\WhatsApp\\WhatsApp.exe"
        try:
            subprocess.Popen([whatsapp_path])
            self.display_async(f"{self.ai_name}: WhatsApp’s up—start chatting!", "ai")
        except FileNotFoundError:
            self.display_async(f"{self.ai_name}: WhatsApp not found at '{whatsapp_path}'—check where it’s at!", "ai")
        except Exception as e:
            self.display_async(f"{self.ai_name}: WhatsApp problem: {e}.", "ai")

    def update_time(self):
        """Updates the time display label every second."""
        current_time = datetime.now().strftime("%I:%M:%S %p")
        self.time_label.config(text=current_time)
        self.master.after(1000, self.update_time)  # Schedule next update in 1 second

    def update_date(self):
        """Updates the date display label once every 24 hours."""
        current_date = datetime.now().strftime("%A, %B %d, %Y")
        self.date_label.config(text=current_date)
        self.master.after(86400000, self.update_date)  # Schedule next update in 24 hours (86,400,000 ms)

    def create_mini_window(self):
        """
        Creates a small, always-on-top control window for quick system actions.
        This window has no title bar and is positioned in the bottom-right.
        """
        if self.mini_controls_window and self.mini_controls_window.winfo_exists():
            self.mini_controls_window.lift()  # Bring to front if already open
            return

        self.mini_controls_window = tk.Toplevel(self.master, bg=BG_DARK)
        self.mini_controls_window.title("Controls")  # Title for internal reference, won't be visible
        self.mini_controls_window.overrideredirect(True)  # Removes window decorations (title bar, borders)
        self.mini_controls_window.attributes("-topmost", True)  # Keep window on top

        # Calculate position for bottom-right corner
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        window_width = 250
        window_height = 180
        x_position = screen_width - window_width - 10  # 10 pixels from right edge
        y_position = screen_height - window_height - 50  # 50 pixels from bottom (to account for taskbar)
        self.mini_controls_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Add a small 'X' button to close the mini window
        close_button = Button(self.mini_controls_window, text="X", command=self.mini_controls_window.destroy,
                              bg="red", fg="white", font=("Arial", 8, "bold"), width=3, height=1)
        close_button.pack(anchor="ne", padx=2, pady=2)  # Position in top-right corner

        # Frame to hold the control buttons
        button_frame = tk.Frame(self.mini_controls_window, bg=BG_DARK)
        button_frame.pack(pady=5)

        # List of control buttons and their corresponding commands
        buttons = [
            ("🔊+", self.volume_up), ("🔊-", self.volume_down), ("🔇 Mute", self.mute_volume),
            ("🎤 Start", self.start_listening), ("⏹ Stop", self.stop_listening),
            ("⏮ Prev", self.previous_track), ("⏯ Play", self.play_pause), ("⏭ Next", self.next_track),
            ("🔆+", self.brightness_up), ("🔆-", self.brightness_down)
        ]

        # Arrange buttons in a grid within the button_frame
        row = 0
        col = 0
        for text, command in buttons:
            Button(button_frame, text=text, command=command, width=8, bg=BG_MEDIUM, fg=TEXT_PRIMARY,
                   font=("Arial", 9), activebackground=BG_MEDIUM).grid(row=row, column=col, padx=3, pady=3)
            col += 1
            if col > 2:  # Max 3 buttons per row
                col = 0
                row += 1


# --- Main Application Class (Tabbed Interface) ---
class MainApp:
    def __init__(self, root):
        """
        Initializes the main Multi-AI Utility application with a tabbed interface.
        :param root: The root Tkinter window.
        """
        self.root = root
        self.root.title("Multi-AI Utility")  # Set the title for the main window
        self.root.geometry("1000x800")  # Set a default window size
        self.root.configure(bg=BG_DARK)  # Set background color for the main window

        # Ensure base directories for memory and models exist
        os.makedirs(MEMORY_BASE_DIR, exist_ok=True)
        os.makedirs(MODEL_BASE_DIR, exist_ok=True)

        # Configure ttk styles for the Notebook (tabs)
        style = ttk.Style()
        style.theme_use("clam")  # A good base theme
        style.configure("TNotebook", background=BG_DARK, borderwidth=0)
        style.configure("TNotebook.Tab", background=BG_MEDIUM, foreground=TEXT_PRIMARY,
                        font=("Arial", 10, "bold"), padding=[10, 5], borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", BG_DARK)],  # Selected tab background
                  foreground=[("selected", TEXT_PRIMARY)])  # Selected tab text color

        # Create a Notebook widget (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both")  # Make notebook fill the main window

        # --- JARVIS Chat AI Tab ---
        self.jarvis_frame = tk.Frame(self.notebook, bg=BG_DARK)  # Frame for JARVIS
        self.notebook.add(self.jarvis_frame, text="🧠 JARVIS AI")  # Add as a tab
        self.jarvis_app = ChatApp(self.jarvis_frame, ai_name="JARVIS")  # Instantiate JARVIS ChatApp

        # Print a launch message to the console
        print("Multi-AI Utility Launched Successfully!")


if __name__ == "__main__":
    # Create the root Tkinter window
    root = tk.Tk()
    # Instantiate the main application
    app = MainApp(root)
    # Start the Tkinter event loop
    root.mainloop()

