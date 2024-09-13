from src.kidase_creator import KidaseCreator

# Example usage with Geez, Tigrinya, and English languages
if __name__ == '__main__':
    kidase_creator = KidaseCreator('./data', ['ግእዝ', 'ትግርኛ', 'english'])
    kidase_creator.create_presentation()
