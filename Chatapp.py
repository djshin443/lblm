import tkinter as tk
from tkinter import messagebox
import openai
import threading
import configparser
from tkinter import filedialog
import os
import sys

# Create a new ConfigParser object
config = configparser.ConfigParser()

# Get the path of the .exe file
exe_path = os.path.dirname(sys.executable)

# Read the config file
config_file_path = os.path.join(exe_path, 'chatapp.conf')
if os.path.exists(config_file_path):
    config.read(config_file_path)
else:
    messagebox.showerror("구성 오류", "chatapp.conf 파일을 찾을 수 없습니다. 새로운 파일이 생성될 것입니다.")
    with open(config_file_path, 'w') as config_file:
        config_file.write('[openai]\napi_key = 실제_openai_api_키를_여기에_입력해주세요')
    config.read(config_file_path)

# Get the API key
try:
    openai.api_key = config.get('openai', 'api_key')
    if openai.api_key.strip() == "" or openai.api_key == "실제_openai_api_키를_여기에_입력해주세요":
        raise ValueError('API key not found or invalid.')
except ValueError as e:
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showerror("구성 오류", "chatapp.conf에서 API 키를 찾을 수 없거나 잘못되었습니다. 실제 OpenAI API 키를 입력하세요.")
    messagebox.showinfo("종료", "API 키를 입력하고 애플리케이션을 다시 시작하세요.")
    root.destroy()  # Destroy the main window
    exit(1)


class ChatApp:
    def __init__(self, root):
        self.root = root
        self.dark_mode = tk.BooleanVar()
        self.message_history = []  # Store the history of messages exchanged
        self.conversation_history = [{"role": "system", "content": "You are a helpful assistant."}]  # Conversation history to be passed to the model

        # Dark mode and light mode settings
        self.dark_mode_colors = {"bg": "#282c34", "fg": "#ffffff"}
        self.light_mode_colors = {"bg": "#ffffff", "fg": "#000000"}

        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=1)

        # Text area for displaying messages
        self.text_area = tk.Text(self.main_frame, wrap="word")
        self.text_area.pack(fill=tk.BOTH, expand=1)

        # Scrollbar for the text area
        self.scrollbar = tk.Scrollbar(self.text_area)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_area.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.text_area.yview)

        # Message input area
        self.entry_frame = tk.Frame(self.main_frame, bg=self.light_mode_colors["bg"])
        self.entry_frame.pack(fill=tk.X, padx=5, pady=5)
        self.message_entry = tk.Text(self.entry_frame, height=5, borderwidth=2, relief="groove")  # Add border and relief style
        self.message_entry.pack(fill=tk.X, side=tk.LEFT, expand=1, padx=5)
        self.message_entry.bind('<Return>', self.send_message)
        self.message_entry.bind('<Shift-Return>', self.new_line)

        self.send_button = tk.Button(self.entry_frame, text="Send", command=self.send_message, relief="groove", borderwidth=2)
        self.send_button.pack(side=tk.RIGHT, padx=5)

        # Dark mode and Light mode toggle button
        self.mode_button = tk.Checkbutton(self.main_frame, text="Dark Mode", variable=self.dark_mode, command=self.switch_mode)
        self.mode_button.pack(side=tk.BOTTOM)

        # Message save button
        self.save_button = tk.Button(self.main_frame, text="Save Messages", command=self.save_message_history)
        self.save_button.pack(side=tk.BOTTOM)

        # Set initial mode
        self.switch_mode()

    def switch_mode(self):
        mode_colors = self.dark_mode_colors if self.dark_mode.get() else self.light_mode_colors
        self.main_frame.configure(bg=mode_colors["bg"])
        self.text_area.configure(bg=mode_colors["bg"], fg=mode_colors["fg"])
        self.entry_frame.configure(bg=mode_colors["bg"])
        self.message_entry.configure(bg=mode_colors["bg"], fg=mode_colors["fg"])
        self.send_button.configure(bg=mode_colors["bg"], fg=mode_colors["fg"], activebackground=mode_colors["fg"], activeforeground=mode_colors["bg"])
        self.mode_button.configure(bg=mode_colors["bg"], fg=mode_colors["fg"], selectcolor=mode_colors["fg"])
        self.save_button.configure(bg=mode_colors["bg"], fg=mode_colors["fg"])

        # Change message colors
        message_color = "#000000" if self.dark_mode.get() else "blue"
        assistant_color = "#FFA500" if self.dark_mode.get() else "green"
        you_color = "#FFFFFF" if self.dark_mode.get() else "blue"
        self.text_area.tag_config("message", foreground=message_color)
        self.text_area.tag_config("assistant", foreground=assistant_color)
        self.text_area.tag_config("you", foreground=you_color)

    def count_tokens(self, text):
        return len(text.split())

    def send_message(self, event=None):
        message = self.message_entry.get("1.0", tk.END).strip()  # Retrieve the message from the text widget
        if not message:
            return

        self.display_message("You: " + message, "you")
        self.conversation_history.append({"role": "user", "content": message})  # Add user message to conversation history
        self.message_entry.delete("1.0", tk.END)
        self.display_message("Assistant is typing...", "gray")

        threading.Thread(target=self.get_assistant_message, args=(message,)).start()

    def new_line(self, event=None):
        self.message_entry.insert(tk.END, '\n')

    def display_message(self, message, tag):
        self.text_area.insert(tk.END, message + "\n", tag)
        self.text_area.see(tk.END)
        self.message_history.append(message + "\n")

    def get_assistant_message(self, message):
        self.conversation_history.append({"role": "assistant", "content": "대답은 한국어로 부탁드립니다."}) # Add the prompt for Korean response
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=self.conversation_history  # Pass conversation history to the model
        )

        assistant_message = response.choices[0].message['content']

        # Add assistant message to conversation history
        self.conversation_history.append({"role": "assistant", "content": assistant_message})

        # If the conversation history exceeds the token limit, remove the oldest message
        while self.count_tokens('\n'.join([m['content'] for m in self.conversation_history])) > 4096:
            self.conversation_history.pop(0)

        self.root.after(0, self.update_assistant_message, assistant_message)  # Call the update method using root's after

    def save_message_history(self):
        file_name = filedialog.asksaveasfilename(initialdir="/", title="Save file",
                                                 filetypes=(("Text Files", "*.txt"), ("all files", "*.*")))

        if not file_name.endswith(".txt"):
            file_name += ".txt"

        with open(file_name, 'w') as file:
            file.writelines(self.message_history)

    def update_assistant_message(self, assistant_message):
        self.text_area.delete('end-2l linestart', 'end-1l lineend')
        self.display_message("Assistant: " + assistant_message, "assistant")

root = tk.Tk()
root.title("Chat with AI")
app = ChatApp(root)
root.mainloop()
