import os

def Speak(text):

    # writing the text to the vbs file.
    with open("command.vbs", "a") as file_h:
        file_h.write(f' "{text}"')

    command = " cscript command.vbs"
    os.system(command)


if __name__ == '__main__':
    print("Welcome to the text-to-speech app.")
    print("#Created by Bhuwan Khatiwada.")

    while(True):

        text = input("Enter what you want to say: ")
        # using command prompt and vbs file.


        # writing just the initial command without the text on every loop.
        cmd='Set speak = CreateObject("SAPI.SpVoice")\nspeak.Speak'
        with open("command.vbs", "w") as file_h:
            file_h.write(cmd)

        if (text == 'quit'):
            Speak("Thank you! Bye-Bye!    ")
            break
        else:
            Speak(text)


        print("Say 'quit' if you want to stop.")




