from win32com.client import Dispatch
import random
def hangman():
    word=random.choice(['python','rider','provider','binod','gajodar','elon','lloyd','vishal','vijesh','vyas'])
    valid_letters='abcdefghijklmnopqrstuvwxyz'
    # turns left
    turns=10
    # guess made by user
    guessmade=""
    while len(word)>0:
#         if generated word lenght is greater than 0,can use try catch also

          main=""
#         generated word if the user input contains in generated word

          # for a letter if its in generated string  add it to main string and display it to user asking the input,else empty string
          for letter in word:
             if letter in guessmade:
                main += letter
             else:
                main += "_" + " "

          guess=input(f"Guess the word {main}")
          turns-=1

          # checking whether the input is alphabet or not
          if guess in valid_letters:
            guessmade += guess
          else:
            print("Enter valid character")
            guess = input()

          # winning condtion
          if main==word:
              speak = Dispatch("SAPI.SpVoice")
              speak.speak("You saved charlse")
              break

          if guess not in word:
              if turns==9:
                  print("9 Chance left")
                  print(" --------- ")
              if turns==8:
                  print("8 chance left")
                  print(" --------- ")
                  print("     O     ")
              if turns==7:
                  print("7 chance left")
                  print(" --------- ")
                  print("     O     ")
                  print("     |     ")
              if turns == 6:
                  print("6 chance left")
                  print(" --------- ")
                  print("     O     ")
                  print("     |     ")
                  print("    /      ")
              if turns == 5:
                  print("5 chance left")
                  print(" --------- ")
                  print("     O     ")
                  print("     |     ")
                  print("    / \    ")
              if turns == 4:
                  print("4 chance left")
                  print(" --------- ")
                  print("   \ O     ")
                  print("     |     ")
                  print("    / \    ")
              if turns == 3:
                  print("3  chance left")
                  print(" --------- ")
                  print("   \ O /   ")
                  print("     |     ")
                  print("    / \    ")
              if turns == 2:
                  print("2 chance left")
                  print(" --------- ")
                  print("     |     ")
                  print("   \ O /   ")
                  print("     |     ")
                  print("    / \    ")
              if turns == 1:
                  print("Last chance ")
                  print(" --------- ")
                  print("     |     ")
                  print("     O     ")
                  print("    /|\    ")
                  print("    / \    ")
              if turns == 0:
                  speak = Dispatch("SAPI.SpVoice")
                  speak.speak("Hogbitaa charlse hogbittaa")
                  print(" --------- ")
                  print("       |    ")
                  print("     O_|   ")
                  print("    /|\     ")
                  print("    / \    ")
                  break



if __name__ == '__main__':
    speak=Dispatch("SAPI.SpVoice")
    name=input("Enter your name:")
    speak.speak(f"Welcome {name}\n Kim has given punishment to charlse \n Guess the word!! only you can save him \nyou have 10 chance")
    hangman()
