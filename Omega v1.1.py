import os
import pyttsx3
import pyautogui
import sys
import asyncio
import re
from pathlib import Path
from dotenv import load_dotenv
from browser_use import Agent, BrowserSession
import customtkinter
import speech_recognition as sr
import threading
from google.oauth2 import service_account
import google.genai as genai
from browser_use.llm import ChatGoogle
import subprocess
import keyboard
from collections import deque



def resource_path(filename):
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


load_dotenv(resource_path(".env"))

# Lazy initialization for faster startup
_main_browser = None
_genai_client = None

def get_browser_session():
    global _main_browser
    if _main_browser is None:
        _main_browser = BrowserSession(
            executable_path='C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            user_data_dir='~/.config/browseruse/profiles/default',
            keep_alive=True,
            highlight_elements=False
        )
    return _main_browser

def get_genai_client():
    global _genai_client
    if _genai_client is None:
        api_key = os.getenv("GOOGLE_API_KEY")
        if not api_key:
            raise ValueError("GOOGLE_API_KEY environment variable not set")
        _genai_client = genai.Client(api_key=api_key)
    return _genai_client

DESKTOP_PATH = os.path.join(os.path.expanduser('~'), 'Desktop')

# Cache system instruction to avoid repeated formatting
_SYSTEM_INSTRUCTION = None

def get_system_instruction():
    global _SYSTEM_INSTRUCTION
    if _SYSTEM_INSTRUCTION is None:
        _SYSTEM_INSTRUCTION = f"""
You are O.M.E.G.A., a smart, voice-powered desktop assistant running locally.
The following modules are ALREADY IMPORTED and available in the global scope:
{', '.join(['os', 'asyncio', 'pyttsx3', 'subprocess', 'pyautogui', 'keyboard', 're'])}.
Do NOT include import statements for these pre-imported modules.
You also have access to a variable `DESKTOP_PATH` which points to the user's Desktop directory (e.g., '{DESKTOP_PATH}').
For other standard Python libraries (e.g., 'time', 'datetime', 'json'), you CAN import them.

Do not even tell what programming language you are using, just give code that can be executed locally.

Always try to perform commands by generating good, safe Python code that controls the desktop. And don't be lazy by generating very small code, generate the code which an innocnet user will be happy with the results.
Examples: open apps like Word, Notepad, Calculator, click positions, type text, etc.
you can use os.startfile() to open them. REMEMBER, it is same as running in "Run".

Use the most efficient code possible like if user asked to open an app, you can use os module or similar. You can code however you like but simple and effective which should work on low-end devices as well.
And always add exception because if user asked to open the app which doesn't exist in the computer and you generate a code to open it, an error will occur.

If the user asks to open an office applications like PowerPoint, Word, or Excel, and asked you to do some task, use python-docx, python-pptx, or openpyxl modules to create the files. You must create the file in the user's Desktop directory using the `DESKTOP_PATH` variable
After creating the file, you must open the file using the `os.startfile()` function with the file path.
If the user asks to open Notepad, create a text file on the Desktop with a meaningful filename based on the context or content (e.g., 'shopping_list.txt', 'notes_2024.txt', 'todo.txt') and open it using `os.startfile()`. Be creative with filenames!

If you need to tell something to the user, use the `speak` function.
The `speak` function is already defined and uses the pyttsx3 library to convert text to speech.

If user asked to navigate the system files, don't search with the accurate file name,
but rather use a general search term like "search for files related to <your topic or keyword>".

IMPORTANT: You must choose ONE of these response types:
1. Generate Python code to handle the task locally
2. OR respond ONLY with "USE_BROWSER_AGENT: [detailed task description]" for web-based tasks

If you cannot solve it using Python code for local desktop control (like specific Google searches, downloading files from the internet, or complex web interactions), respond ONLY with:
"USE_BROWSER_AGENT: [detailed task description]"
where [detailed task description] is a clear, specific description of what the browser agent should accomplish. Example: If user asks to download something, don't just tell the browser use, "download the thing which user is asking for.". Instead be more clear what the user the asking by looking at the previous conversation.

NEVER mix Python code with browser agent instructions in the same response.
Never delete files (except temporary files you create),
shut down the system, or do anything dangerous.
Keep code minimal, efficient, and specific to the task.
Always try to execute code directly unless it's totally not executable locally.
Use subprocess.Popen() or subprocess.run() instead of os.system() or subprocess.call() for better security.

If user say "Omega exit", "Omega quit", "bye", "goodbye", or similar exit commands, call the exit_app() function to close the program gracefully.

And finally, you are developed by a class 12th high school student from International Indian school, Jeddah. Name: Mohamed Nadeem.
"""
    return _SYSTEM_INSTRUCTION

PRE_IMPORTED_MODULES = {"os", "asyncio", "pyttsx3", "subprocess", "pyautogui", "keyboard", "re"}

# Removed global TTS engine to avoid threading conflicts

# Thread lock for TTS to prevent concurrent access
_tts_lock = threading.Lock()


def speak(text, output_box=None):
    # Always show the text immediately
    if output_box is not None:
        output_box.configure(state="normal")
        output_box.insert("end", f"Omega: {text}\n")
        output_box.configure(state="disabled")
        output_box.see("end")
        output_box.update_idletasks()  # Ensure the text appears before speaking
    else:
        print("Omega:", text)
    
    # Run TTS in background thread with proper synchronization
    def tts_thread():
        with _tts_lock:  # Ensure only one TTS operation at a time
            try:
                # Create a fresh engine instance each time to avoid conflicts
                engine = pyttsx3.init()
                engine.setProperty('rate', 200)
                engine.setProperty('volume', 0.9)
                engine.say(text)
                engine.runAndWait()
                engine.stop()  # Clean up
            except Exception as e:
                print(f"TTS Error: {e}")
                # Fallback: try with system default settings
                try:
                    fallback_engine = pyttsx3.init()
                    fallback_engine.say(text)
                    fallback_engine.runAndWait()
                    fallback_engine.stop()
                except Exception:
                    print(f"TTS completely failed for: {text}")
    
    threading.Thread(target=tts_thread, daemon=True).start()


# Pre-compile regex patterns for better performance
_SYSTEM_DANGEROUS_PATTERNS = [
    re.compile(r"\bformat\s+[A-Z]:", re.IGNORECASE),  # Format drive commands
    re.compile(r"shutdown\s+/", re.IGNORECASE),       # Windows shutdown commands
    re.compile(r"reboot\s+/", re.IGNORECASE),         # Windows reboot commands
    re.compile(r"rmdir\s+/s", re.IGNORECASE),         # Recursive directory deletion
    re.compile(r"del\s+/[sq]", re.IGNORECASE),        # Force delete commands
    re.compile(r"rm\s+-rf", re.IGNORECASE),           # Unix-style force delete
]

_SAFE_SUBPROCESS_PATTERNS = [
    re.compile(r"subprocess\.Popen\("),  # Allow any subprocess.Popen call
    re.compile(r"subprocess\.run\("),    # Allow any subprocess.run call
]

def is_safe_code(code_str):
    # Critical security patterns that should never be allowed
    critical_dangerous_patterns = [
        "eval(", "exec(",  # Code injection
        "ctypes.windll",   # Direct Windows API access
        "socket.",         # Network operations
    ]
    
    # Check critical patterns first
    for pattern in critical_dangerous_patterns:
        if pattern in code_str:
            return False
    
    # Check system patterns with pre-compiled regex
    for pattern in _SYSTEM_DANGEROUS_PATTERNS:
        if pattern.search(code_str):
            return False
    
    # Additional checks for subprocess usage
    if "subprocess." in code_str:
        # Allow common safe subprocess operations
        if any(pattern.search(code_str) for pattern in _SAFE_SUBPROCESS_PATTERNS):
            return True
        else:
            # Block subprocess usage that doesn't match safe patterns
            return False
    
    return True  # Default to allowing code unless explicitly dangerous


def strip_unnecessary_imports(code_str):
    lines = code_str.splitlines()
    filtered_lines = []
    for line in lines:
        stripped_line = line.strip()
        is_import_line = stripped_line.startswith("import ") or stripped_line.startswith("from ")
        
        if is_import_line:
            try:
                parts = stripped_line.split()
                module_name_to_check = parts[1].split('.')[0] 
                if module_name_to_check in PRE_IMPORTED_MODULES:
                    print(f"O.M.E.G.A. (Code Util): Stripping unnecessary import: {line}")
                    continue 
            except IndexError:
                pass
        filtered_lines.append(line)
    return "\n".join(filtered_lines)


from collections import deque

CONVERSATION_HISTORY = deque(maxlen=10)  # Remember last 10 conversations

def run_ai_command(command, output_box=None):
    try:
        os.makedirs(DESKTOP_PATH, exist_ok=True)
    except Exception as e:
        print(f"O.M.E.G.A. (Setup Error): Could not create Desktop path {DESKTOP_PATH}: {e}")
    
    # Build history prompt more efficiently
    history_prompt = ""
    if CONVERSATION_HISTORY:
        history_prompt = "\n--- RECENT CONVERSATION HISTORY ---\n"
        for i, (user_cmd, ai_resp) in enumerate(CONVERSATION_HISTORY, 1):
            history_prompt += f"[{i}] User: {user_cmd}\n[{i}] Omega: {ai_resp}\n\n"
        history_prompt += "--- END HISTORY ---\n"
    
    prompt = (
        f"{get_system_instruction()}\n"
        f"{history_prompt}\n"
        f"Command: {command}\nGenerate Python code:"
    )

    print("O.M.E.G.A. (AI): Querying LLM...")
    client = get_genai_client()
    response = client.models.generate_content(model="gemini-flash-latest", contents=prompt,)
    
    generated_text = response.candidates[0].content.parts[0].text.strip()
    
    if generated_text.startswith("USE_BROWSER_AGENT"):
        print("O.M.E.G.A. (AI): LLM requested browser agent.")
        # Extract custom task description if provided
        if ":" in generated_text:
            browser_task = generated_text.split(":", 1)[1].strip()
        else:
            browser_task = command  # Fallback to original command
        CONVERSATION_HISTORY.append((command, generated_text))
        return "USE_BROWSER_AGENT", browser_task

    code_to_run = generated_text.strip('`')
    if code_to_run.lstrip().startswith("python"):
        code_to_run = code_to_run.lstrip()[len("python"):].lstrip()
    code_to_run = strip_unnecessary_imports(code_to_run)

    if not is_safe_code(code_to_run):
        print(f"O.M.E.G.A. (Safety): Detected potentially unsafe code. Execution aborted.\n{code_to_run}")
        CONVERSATION_HISTORY.append((command, "[UNSAFE_CODE]"))
        return "UNSAFE_CODE", None

    if not code_to_run:
        print("O.M.E.G.A. (AI): LLM generated empty code.")
        CONVERSATION_HISTORY.append((command, "[EMPTY_CODE]"))
        return "EMPTY_CODE", None

    try:
        # Capture what the user sees
        captured_output = []
        
        exec_globals = {k: v for k, v in globals().items() if k in PRE_IMPORTED_MODULES or k in ['speak', 'print', 'DESKTOP_PATH']}
        exec_globals['output_box'] = output_box
        
        # Enhanced speak function that captures what was spoken
        def tracked_speak(text):
            captured_output.append(f"Spoke: {text}")
            speak(text, output_box)
        exec_globals['speak'] = tracked_speak
        
        # Enhanced print function that captures console output
        def tracked_print(*args, **kwargs):
            output_text = " ".join(str(arg) for arg in args)
            captured_output.append(f"Printed: {output_text}")
            print(*args, **kwargs)
        exec_globals['print'] = tracked_print
        
        # Add exit functionality for AI
        def safe_exit_app():
            captured_output.append("Action: Application exit requested")
            if output_box:
                # Find the root window safely
                widget = output_box
                while widget and hasattr(widget, 'master') and widget.master:
                    widget = widget.master
                if widget and hasattr(widget, 'after'):
                    widget.after(1000, widget.destroy)
            else:
                sys.exit()
        exec_globals['exit_app'] = safe_exit_app
        
        print("O.M.E.G.A. (Exec): Attempting to execute code.")
        exec(code_to_run, exec_globals)
        
        # Build comprehensive result for conversation history
        execution_result = f"Code: {code_to_run}"
        if captured_output:
            execution_result += f"\nUser saw: {'; '.join(captured_output)}"
        else:
            execution_result += "\nUser saw: Code executed silently (no visible output)"
            
        CONVERSATION_HISTORY.append((command, execution_result))
        return "SUCCESS", "Done"
    except Exception as e:
        error_msg = f"Code: {code_to_run}\nUser saw: ERROR - {str(e)}"
        print(f"O.M.E.G.A. (Exec Error): {e}\nCode that failed:\n{code_to_run}")
        CONVERSATION_HISTORY.append((command, error_msg))
        return "EXECUTION_ERROR", f"Execution error: {e}"

async def run_browser_task(task_description):
    print(f"O.M.E.G.A. (Browser): Initializing browser agent for task: {task_description}")
    model = ChatGoogle(model='gemini-2.5-flash')
    browser_session = get_browser_session()
    agent = Agent(task=task_description, llm=model, browser_session=browser_session)
    try:
        result = await agent.run()
        return str(result.final_result()) if result else "Browser task completed with no specific result."
    except Exception as e:
        print(f"O.M.E.G.A. (Browser Error): {e}")
        return f"Browser task error: {e}"

class OmegaDesktopAssistant(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title("O.M.E.G.A. Desktop Assistant")
        self.geometry("700x550")
        
        # Set icon if available
        try:
            self.iconbitmap(resource_path("logo.ico"))
        except:
            pass

        # Set overall theme
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("blue")

        # Main frame
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=15)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Chat display area
        self.chat_display = customtkinter.CTkTextbox(self.main_frame, width=600, height=350,
                                                      corner_radius=10,
                                                      font=("Arial", 14),
                                                      wrap="word")
        self.chat_display.pack(pady=10, padx=10, fill="both", expand=True)
        self.chat_display.insert("0.0", "Omega: How can I help you?\n")
        self.chat_display.configure(state="disabled")

        # Input frame
        self.input_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=10)
        self.input_frame.pack(pady=10, padx=10, fill="x")

        # Text input field
        self.user_input = customtkinter.CTkEntry(self.input_frame,
                                                 placeholder_text="Type your message here...",
                                                 width=400,
                                                 height=40,
                                                 border_width=2,
                                                 corner_radius=10,
                                                 font=("Arial", 12))
        self.user_input.pack(side="left", pady=10, padx=10, fill="x", expand=True)
        self.user_input.focus()

        # Send button
        self.send_button = customtkinter.CTkButton(self.input_frame, text="Send",
                                                   command=self.send_command,
                                                   width=80, height=40,
                                                   corner_radius=10,
                                                   font=("Arial", 12, "bold"))
        self.send_button.pack(side="left", pady=10, padx=(0, 5))

        # Speak button
        self.speak_button = customtkinter.CTkButton(self.input_frame, text="Speak",
                                                    command=self.recognize_voice,
                                                    width=80, height=40,
                                                    corner_radius=10,
                                                    font=("Arial", 12, "bold"))
        self.speak_button.pack(side="left", pady=10, padx=(0, 10))

        # Bind Enter key to send command only when input field has focus
        self.user_input.bind('<Return>', lambda event: self.send_command())
        
        # Handle window close event properly
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def run_browser_task_thread(self, browser_task, original_command=None):
        def task():
            result = asyncio.run(run_browser_task(browser_task))
            # Update conversation history with browser result
            if original_command:
                browser_result = f"Browser Task: {browser_task}\nUser saw: {result}"
                # Update the last conversation entry with the actual result
                if CONVERSATION_HISTORY and CONVERSATION_HISTORY[-1][0] == original_command:
                    CONVERSATION_HISTORY[-1] = (original_command, browser_result)
            self.after(0, lambda: self.insert_message(f"Omega: {result}"))
        threading.Thread(target=task, daemon=True).start()

    def insert_message(self, message):
        self.chat_display.configure(state="normal")
        self.chat_display.insert("end", f"{message}\n")
        self.chat_display.configure(state="disabled")
        self.chat_display.see("end")

    def send_command(self):
        command = self.user_input.get()
        if not command.strip():
            return
        
        self.insert_message(f"You: {command}")
        self.user_input.delete(0, "end")
        
        # Check for exit commands before processing
        def exit_app():
            self.insert_message("Omega: Goodbye!")
            self.after(1500, self.destroy)  # Close after 1.5 second
            return
        
        
        # Show processing indicator
        self.send_button.configure(text="Processing...", state="disabled")
        self.speak_button.configure(state="disabled")
        
        # Process AI command in background thread
        def process_command():
            try:
                status, ai_result_msg = run_ai_command(command, output_box=self.chat_display)
                # Update UI on main thread
                self.after(0, lambda: self.handle_ai_response(status, ai_result_msg, command))
            except Exception as e:
                error_msg = str(e)  # Capture the error message immediately
                self.after(0, lambda: self.handle_ai_response("ERROR", error_msg, command))

        threading.Thread(target=process_command, daemon=True).start()

    
    def handle_ai_response(self, status, ai_result_msg, command):
        # Re-enable buttons
        self.send_button.configure(text="Send", state="normal")
        self.speak_button.configure(state="normal")
        
        if status == "SUCCESS":
            if ai_result_msg and ai_result_msg.strip().lower() != "done":
                self.insert_message(f"Omega: {ai_result_msg}")
        elif status == "USE_BROWSER_AGENT":
            self.insert_message("Omega: Switching to browser agent for this task.")
            # ai_result_msg now contains the custom browser task description
            browser_task = ai_result_msg if ai_result_msg else command
            self.run_browser_task_thread(browser_task, command)
        elif status == "UNSAFE_CODE":
            self.insert_message("Omega: An error has occured, please try again")
        elif status == "EXECUTION_ERROR":
            self.insert_message("Omega: An error has occured, please try again")
        elif status == "EMPTY_CODE":
            self.insert_message("Omega: An error has occured, please check you're wifi")
        else:
            self.insert_message("Omega: An error has occured, please try again")
            self.run_browser_task_thread(command, command)

    def recognize_voice(self):
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            self.insert_message("Omega: Listening...")
            self.update()
            try:
                audio = recognizer.listen(source, timeout=5)
                command = recognizer.recognize_google(audio)
                self.user_input.delete(0, "end")
                self.user_input.insert(0, command)
                self.insert_message(f"Omega: Heard '{command}'")
            except sr.WaitTimeoutError:
                self.insert_message("Omega: Listening timed out. Please try again.")
            except sr.UnknownValueError:
                self.insert_message("Omega: Could not understand audio.")
            except sr.RequestError as e:
                self.insert_message(f"Omega: Could not request results; {e}")

    def on_closing(self):
        """Handle window close event with confirmation"""
        import tkinter.messagebox as msgbox
        if msgbox.askokcancel("Exit O.M.E.G.A.", "Are you sure you want to exit O.M.E.G.A.?"):
            try:
                # Clean up browser session if needed
                global _main_browser, _tts_engine
                if _main_browser is not None:
                    _main_browser.close()
                if _tts_engine is not None:
                    _tts_engine.stop()
            except:
                pass
            self.destroy()

def Omega_gui():
    app = OmegaDesktopAssistant()
    app.mainloop()

if __name__ == "__main__":
    Omega_gui()
