import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pytube import YouTube
import threading

# Creating the root window
root = tk.Tk()
root.title("YouTube Video Downloader ")
root.geometry("600x300")
root.resizable(False, False)
root.config(bg="lightblue")

# Creating the StringVar variables
video_Link = tk.StringVar()
download_Path = tk.StringVar()

# Creating the Widgets function
def Widgets():
    # Creating the title label
    title_label = tk.Label(root, text="YouTube Video Downloader", font="SegoeUI 16 bold", bg="lightblue", fg="red")
    title_label.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

    # Creating the link label
    link_label = tk.Label(root, text="YouTube link:", font="Arial 14", bg="lightblue", fg="black")
    link_label.grid(row=1, column=0, padx=10, pady=10, sticky="E")

    # Creating the link entry
    link_entry = tk.Entry(root, width=40, textvariable=video_Link, font="Arial 14")
    link_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=10)

    # Creating the destination label
    destination_label = tk.Label(root, text="Destination:", font="Arial 14", bg="lightblue", fg="black")
    destination_label.grid(row=2, column=0, padx=10, pady=10, sticky="E")

    # Creating the destination entry
    destination_entry = tk.Entry(root, width=30, textvariable=download_Path, font="Arial 14")
    destination_entry.grid(row=2, column=1, padx=10, pady=10)

    # Creating the browse button
    browse_button = tk.Button(root, text="Browse", command=Browse, width=10, font="Arial 12", bg="bisque", relief="groove")
    browse_button.grid(row=2, column=2, padx=10, pady=10)

    # Creating the download button
    download_button = tk.Button(root, text="Download Video", command=Download, width=20, font="Arial 12", bg="thistle1", relief="groove")
    download_button.grid(row=3, column=1, padx=10, pady=10)

    # Creating the progress bar
    global progress_bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress_bar.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

# Creating the Browse function
def Browse():
    # Asking the user to select a destination folder
    download_Directory = filedialog.askdirectory(initialdir="YOUR DIRECTORY PATH", title="Save Video")
    # Setting the download_Path variable to the selected folder
    download_Path.set(download_Directory)

# Creating the Download function
def Download():
    # Starting a new thread to perform the download in the background
    download_thread = threading.Thread(target=DownloadVideo)
    download_thread.start()

# Function to download the video
def DownloadVideo():
    # Getting the YouTube link from the link entry
    Youtube_link = video_Link.get()
    # Getting the download folder from the destination entry
    download_Folder = download_Path.get()

    # Validating if link and folder are not empty
    if not Youtube_link:
        messagebox.showerror("Error", "Please enter a valid YouTube link")
        return

    if not download_Folder:
        messagebox.showerror("Error", "Please select a download destination")
        return

    try:
        # Creating an object of YouTube class
        getVideo = YouTube(Youtube_link, on_progress_callback=progress_function)
        # Getting the highest resolution stream of the video
        videoStream = getVideo.streams.get_highest_resolution()
        # Starting the download
        videoStream.download(download_Folder)
        # Displaying a success message
        messagebox.showinfo("SUCCESS", f"DOWNLOADED AND SAVED IN\n{download_Folder}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to download video. Reason: {str(e)}")

# Callback function to update progress bar
def progress_function(stream, chunk, remaining):
    percent = int((100 * (stream.filesize - remaining)) / stream.filesize)
    progress_bar["value"] = percent
    root.update_idletasks()  # Make sure to update the GUI

# Calling the Widgets function
Widgets()

# Running the mainloop
root.mainloop()
