import random

while True:
    choices = ["rock", "paper", "scissors"]
    computer = random.choice(choices)
    player = None

    while player not in choices:
        player = input("Rock, Paper or Scissors?: ").lower()

    if player == computer:  
        print("Computer: ", computer)
        print("Player: ", player)
        print("TIE!!!!")
        
    elif player == "rock":
        if computer == "scissors":
            print("Computer: ", computer)
            print("Player: ", player)
            print(" YOU WIN!!!")
        if computer == "paper":
            print("Computer: ", computer)
            print("Player: ", player)
            print("You LOSE!!!!")

    elif player == "paper":
        if computer == "scissors":
            print("Computer: ", computer)
            print("Player: ", player)
            print(" YOU LOSE!!!")
        if computer == "rock":
            print("Computer: ", computer)
            print("Player: ", player)
            print("You WIN!!!!")

    elif player == "scissors":
        if computer == "rock":
            print("Computer: ", computer)
            print("Player: ", player)
            print(" YOU LOSE!!!")
        if computer == "paper":
            print("Computer: ", computer)
            print("Player: ", player)
            print("You WIN!!!!")       
        
    play_again = input("Play Again? (Y/N): ").lower()
    if play_again != "y":
        break
print("BYE!")