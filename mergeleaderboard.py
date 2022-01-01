import pathlib
from pathlib import Path

def find_newly_downloaded_files():
    current_path = Path(".")
    answer_detection = current_path.resolve() / "AnswerDetection"
    excel_files = list(answer_detection.glob("*.xlsx"))
    print(excel_files)
    for file in excel_files:
        print(file.as_posix())
        if("Math Commitee Week" in file.as_posix()):
            weekly_scores = file.as_posix()
        elif ("Leaderboard" in file.as_posix()):
            #file containing the data for the current leaderboard
            weekly_leaderboard = file.as_posix()
    
    return (weekly_scores,weekly_leaderboard)

if __name__ == "__main__":
    find_newly_downloaded_files()
