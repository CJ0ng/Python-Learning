import pyjokes

def tell_joke():
    joke = pyjokes.get_joke()
    print(joke)

def main():
    print("Press Enter to hear a joke. Press 'q' and Enter to stop.")
    while True:
        response = input()
        if response.lower() == 'q':
            break
        tell_joke()
        
        



