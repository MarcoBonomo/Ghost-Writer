"""
Ghost Writer with Outlook Integration
Alternates between Notepad typing and Outlook email drafts.
Pattern: Notepad → Outlook → Notepad → (repeat)
"""

import random
import time
import subprocess
import win32com.client
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController
from pynput import keyboard
import sys

# ============================================================================
# CONFIGURATION
# ============================================================================

WORDS = [
    "the", "quick", "brown", "fox", "jumps", "over", "lazy", "dog", "apple",
    "banana", "cherry", "delta", "echo", "foxtrot", "golf", "hotel", "india",
    "juliet", "kilo", "lima", "mike", "november", "oscar", "papa", "quebec",
    "romeo", "sierra", "tango", "uniform", "victor", "whiskey", "xray", "yankee",
    "zulu", "alpha", "bravo", "charlie", "computer", "keyboard", "mouse", "screen",
    "window", "document", "folder", "file", "system", "network", "internet",
    "browser", "search", "engine", "database", "server", "client", "application",
    "program", "software", "hardware", "memory", "processor", "storage", "cloud",
    "security", "password", "username", "login", "logout", "session", "cookie",
    "cache", "buffer", "queue", "stack", "array", "list", "dictionary", "set",
    "tuple", "string", "integer", "float", "boolean", "variable", "function",
    "method", "class", "object", "instance", "module", "package", "library",
    "framework", "algorithm", "data", "structure", "loop", "condition", "branch",
    "recursion", "iteration", "parameter", "argument", "return", "value", "type",
    "meeting", "project", "deadline", "schedule", "calendar", "appointment", "presentation",
    "report", "analysis", "strategy", "planning", "implementation", "execution", "delivery",
    "stakeholder", "collaboration", "communication", "feedback", "review", "approval", "budget",
    "timeline", "milestone", "objective", "goal", "target", "metric", "performance", "outcome",
    "initiative", "opportunity", "challenge", "solution", "recommendation", "decision", "action",
    "update", "status", "progress", "completion", "achievement", "success", "improvement"
]

# Typing configuration
TYPING_SPEED_MIN = 0.05
TYPING_SPEED_MAX = 0.5
WORD_PAUSE_MIN = 0.2
WORD_PAUSE_MAX = 2.0

# Notepad configuration
NOTEPAD_WORDS_PER_CYCLE = 85  # Single stream of max 85 words

# Outlook configuration
OUTLOOK_DISPLAY_TIME = 8  # Seconds to display email before closing

# Global abort flag
abort_script = False

# ============================================================================
# EMERGENCY ABORT HANDLER
# ============================================================================

def on_press(key):
    """Listen for ESC key to abort script."""
    global abort_script
    if key == keyboard.Key.esc:
        print("\n[!] ESC pressed - Aborting script...")
        abort_script = True
        return False

def start_abort_listener():
    """Start background listener for abort command."""
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    return listener

# ============================================================================
# NOTEPAD MANAGEMENT
# ============================================================================

def open_notepad_if_needed():
    """Open Notepad only if it's not already running."""
    try:
        import psutil
        notepad_running = any('notepad.exe' in p.name().lower() 
                             for p in psutil.process_iter(['name']))
        
        if notepad_running:
            print("[NOTEPAD] Already running - skipping launch")
            return True
        else:
            print("[NOTEPAD] Launching...")
            subprocess.Popen("notepad.exe")
            time.sleep(2)
            print("[NOTEPAD] Opened successfully")
            return True
            
    except ImportError:
        print("[NOTEPAD] Launching (psutil not available)...")
        try:
            subprocess.Popen("notepad.exe")
            time.sleep(2)
            print("[NOTEPAD] Opened successfully")
            return True
        except Exception as e:
            print(f"[NOTEPAD] Failed to open: {e}")
            return False
    except Exception as e:
        print(f"[NOTEPAD] Error: {e}")
        return False

def create_new_notepad_tab(keyboard_ctrl):
    """Create a new tab in Notepad using Ctrl+N."""
    print("[NOTEPAD] Creating new tab (Ctrl+N)...")
    
    keyboard_ctrl.press(Key.ctrl)
    time.sleep(0.05)
    keyboard_ctrl.press('n')
    time.sleep(0.05)
    keyboard_ctrl.release('n')
    keyboard_ctrl.release(Key.ctrl)
    
    time.sleep(1)
    print("[NOTEPAD] New tab created")

# ============================================================================
# OUTLOOK EMAIL GENERATION
# ============================================================================

def generate_professional_email():
    """Generate a professional-looking email with random text."""
    
    sentences = [
        "I hope this message finds you well and that you have had a productive week so far.",
        "Following our recent discussion regarding the quarterly projections, I wanted to share some additional insights.",
        "The team has been working diligently on the strategic initiative we outlined during our planning session.",
        "I am writing to provide an update on the current status of our ongoing projects and key milestones.",
        "Our analysis indicates that implementing these proposed changes could result in improved efficiency.",
        "We have identified several opportunities for optimization that could enhance operational performance.",
        "The preliminary findings from our market research suggest that customer preferences are shifting significantly.",
        "I wanted to reach out to discuss the possibility of scheduling a meeting to review the roadmap.",
        "Based on the feedback we received from various departments, we have refined our approach considerably.",
        "It would be beneficial to collaborate on this initiative as it aligns with our strategic goals.",
        "Please let me know if you have any questions or would like to discuss this proposal further.",
        "I appreciate your continued support and look forward to working together on this project.",
        "The revised timeline accounts for potential challenges and includes contingency plans to mitigate risks.",
        "Our stakeholders have expressed enthusiasm about the direction we are taking with this initiative.",
        "I believe this represents a valuable opportunity to strengthen our competitive position in the market.",
        "Thank you for taking the time to review this information and for your thoughtful consideration.",
        "The data clearly supports the recommendation to proceed with the next phase of implementation.",
        "I am available to answer any questions and can provide additional documentation if needed.",
        "This initiative has the potential to transform how we operate and create lasting benefits.",
        "Your insights and expertise would be invaluable as we navigate this transition successfully."
    ]
    
    # Select random sentences
    num_sentences = random.randint(8, 12)
    selected_sentences = random.sample(sentences, num_sentences)
    
    # Create paragraphs (2-3 sentences each)
    paragraphs = []
    i = 0
    while i < len(selected_sentences):
        paragraph_length = random.randint(2, 3)
        paragraph = " ".join(selected_sentences[i:i + paragraph_length])
        paragraphs.append(paragraph)
        i += paragraph_length
    
    email_body = "\n\n".join(paragraphs)
    
    greetings = ["Dear Team,", "Dear Colleagues,", "Hello,", "Good afternoon,", "Dear All,"]
    closings = ["Best regards,", "Kind regards,", "Sincerely,", "Thank you,", "Warm regards,"]
    
    full_email = f"{random.choice(greetings)}\n\n{email_body}\n\n{random.choice(closings)}\n[Your Name]"
    
    return full_email

def create_and_discard_outlook_draft():
    """Create Outlook email draft and close without saving."""
    try:
        print("[OUTLOOK] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        print("[OUTLOOK] Creating new email draft...")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Generate professional email content
        email_content = generate_professional_email()
        
        # Set email properties
        mail.To = ""
        mail.Subject = "Project Update and Strategic Recommendations"
        mail.Body = email_content
        
        # Display the email
        print("[OUTLOOK] Opening email draft window...")
        mail.Display(False)  # False = non-modal
        
        # Wait so user can see it
        print(f"[OUTLOOK] Email visible for {OUTLOOK_DISPLAY_TIME} seconds...")
        for i in range(OUTLOOK_DISPLAY_TIME):
            if abort_script:
                break
            time.sleep(1)
        
        # Close without saving
        print("[OUTLOOK] Closing draft without saving...")
        inspector = mail.GetInspector
        inspector.Close(1)  # 1 = olDiscard
        
        print("[OUTLOOK] ✓ Draft closed successfully")
        return True
        
    except Exception as e:
        print(f"[OUTLOOK] ✗ Error: {e}")
        return False

# ============================================================================
# TYPING FUNCTIONS
# ============================================================================

def type_word(keyboard_ctrl, word):
    """Type a single word with human-like timing."""
    if abort_script:
        return
    
    for char in word:
        if abort_script:
            return
        keyboard_ctrl.press(char)
        keyboard_ctrl.release(char)
        time.sleep(random.uniform(TYPING_SPEED_MIN, TYPING_SPEED_MAX))

def notepad_typing_cycle(keyboard_ctrl, word_count):
    """Type a stream of words in Notepad."""
    words = random.choices(WORDS, k=word_count)
    
    print(f"[NOTEPAD] Typing {word_count} words...")
    
    for idx, word in enumerate(words):
        if abort_script:
            return False
        
        type_word(keyboard_ctrl, word)
        
        # Add space after word (except last)
        if idx < len(words) - 1:
            keyboard_ctrl.press(Key.space)
            keyboard_ctrl.release(Key.space)
            
            # Pause between words
            pause_time = random.uniform(WORD_PAUSE_MIN, WORD_PAUSE_MAX)
            time.sleep(pause_time)
    
    print(f"[NOTEPAD] ✓ Completed {word_count} words")
    return True

# ============================================================================
# MAIN WORKFLOW
# ============================================================================

def main():
    """Main execution with alternating Notepad → Outlook → Notepad pattern."""
    global abort_script
    
    print("=" * 80)
    print("GHOST WRITER - NOTEPAD & OUTLOOK AUTOMATION")
    print("=" * 80)
    print(f"[CONFIG] Pattern: Notepad → Outlook → Notepad → (repeat)")
    print(f"[CONFIG] Notepad words per cycle: {NOTEPAD_WORDS_PER_CYCLE}")
    print(f"[CONFIG] Outlook display time: {OUTLOOK_DISPLAY_TIME} seconds")
    print("[!] Press ESC at any time to abort")
    print("=" * 80)
    print()
    
    # Start abort listener
    listener = start_abort_listener()
    
    try:
        # Initialize Notepad
        if not open_notepad_if_needed():
            print("[-] Cannot proceed without Notepad")
            return
        
        if abort_script:
            return
        
        # Initialize keyboard controller
        keyboard_ctrl = KeyboardController()
        
        # Cycle counter
        cycle_count = 0
        
        # Main loop: Notepad → Outlook → Notepad → ...
        while not abort_script:
            cycle_count += 1
            
            # ============================================================
            # STEP 1: NOTEPAD (First occurrence in cycle)
            # ============================================================
            print(f"\n{'='*80}")
            print(f"[CYCLE {cycle_count}] STEP 1/3: NOTEPAD")
            print(f"{'='*80}\n")
            
            create_new_notepad_tab(keyboard_ctrl)
            
            if abort_script:
                break
            
            time.sleep(1)
            
            success = notepad_typing_cycle(keyboard_ctrl, NOTEPAD_WORDS_PER_CYCLE)
            if not success:
                break
            
            time.sleep(2)
            
            # ============================================================
            # STEP 2: OUTLOOK
            # ============================================================
            print(f"\n{'='*80}")
            print(f"[CYCLE {cycle_count}] STEP 2/3: OUTLOOK")
            print(f"{'='*80}\n")
            
            if abort_script:
                break
            
            success = create_and_discard_outlook_draft()
            if not success:
                print("[!] Outlook step failed, but continuing...")
            
            time.sleep(2)
            
            # ============================================================
            # STEP 3: NOTEPAD (Second occurrence in cycle)
            # ============================================================
            print(f"\n{'='*80}")
            print(f"[CYCLE {cycle_count}] STEP 3/3: NOTEPAD")
            print(f"{'='*80}\n")
            
            if abort_script:
                break
            
            create_new_notepad_tab(keyboard_ctrl)
            
            if abort_script:
                break
            
            time.sleep(1)
            
            success = notepad_typing_cycle(keyboard_ctrl, NOTEPAD_WORDS_PER_CYCLE)
            if not success:
                break
            
            # End of cycle
            if not abort_script:
                print(f"\n[CYCLE {cycle_count}] ✓ Complete (Notepad → Outlook → Notepad)")
                print("[*] Preparing for next cycle...\n")
                time.sleep(2)
        
        if abort_script:
            print("\n" + "=" * 80)
            print("[!] Script terminated by user")
            print(f"[STATS] Completed {cycle_count} full cycles")
            print("=" * 80)
        
    except Exception as e:
        print(f"\n[-] Error occurred: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if listener.is_alive():
            listener.stop()

# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    main()
