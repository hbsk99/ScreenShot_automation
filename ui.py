from tkinter import *
from tkinter import filedialog, messagebox
from backend import capture_events

def create_ui():
    def start_program():
        capture_events(status_label, stop=False)
        status_label.config(text='Screen events capturing started!')
    
    def stop_program():
        capture_events(status_label, stop=True)
        status_label.config(text='Screen events capturing stopped!')
    
    def update_status_label(text):
        status_label.config(text=text)
    
    # Create the UI window
    root = Tk()
    root.title('Automated SS Capturing TOOL')
    
    # Create the status label
    status_label = Label(root, text='Ready')
    status_label.pack()
    
    # Create the buttons
    start_button = Button(root, text='Start', command=start_program)
    start_button.pack()
    
    stop_button = Button(root, text='Stop', command=stop_program)
    stop_button.pack()
    
    # Start the UI loop
    root.mainloop()

# Run the UI
if __name__ == '__main__':
    create_ui()
    
#modified
