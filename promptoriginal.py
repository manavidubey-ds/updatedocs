import docx
import re
from fuzzywuzzy import fuzz

def preprocess_question(question):
    match = re.search(r'^(.*?)\?', question)
    if match:
        return match.group(1) + "?"
    return question

def extract_and_match(docx_file, prompt_file, output_file="matched_questions.txt", similarity_threshold=90):
    """
    Extracts content from a .docx file and matches questions from a prompt file,
    storing the matched lines in a separate text file.

    Args:
        docx_file (str): Path to the .docx file.
        prompt_file (str): Path to the prompt.txt file.
        output_file (str): Path to the output file to store matched questions.
        similarity_threshold (int): The minimum similarity percentage (0-100).

    Returns:
        list: A list of matched lines from the document.
    """
    matched_lines = []

    try:
        doc = docx.Document(docx_file)
        full_text = "\n".join([para.text for para in doc.paragraphs])
        lines = full_text.split('\n')

        with open(prompt_file, "r", encoding="utf-8") as file:
            questions = [line.strip() for line in file]

        for question in questions:
            best_match = None
            best_similarity = 0
            processed_question = preprocess_question(question)

            for line in lines:
                similarity = fuzz.ratio(processed_question, line)
                if similarity > best_similarity:
                    best_similarity = similarity
                    best_match = line

            if best_similarity >= similarity_threshold:
                matched_lines.append(best_match)

        # Write matched lines to the output file
        with open(output_file, "w", encoding="utf-8") as outfile:
            for line in matched_lines:
                outfile.write(line + "\n")

        print(f"Matched questions saved to: {output_file}")
        return matched_lines

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
        return []
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

# Example Usage
if __name__ == "__main__":
    docx_file_path = "og.docx"
    prompt_file_path = "prompt.txt"
    output_file_path = "matched_questions.txt"  # Specify the output file name

    matched_questions = extract_and_match(docx_file_path, prompt_file_path, output_file_path, similarity_threshold=90)

