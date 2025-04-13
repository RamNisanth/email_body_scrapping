from transformers import pipeline

# Load the summarization pipeline
summarizer = pipeline("summarization", model="facebook/bart-large-cnn")

# Input text
text = """
A good paragraph for summarization should clearly present a single main idea, supported by relevant details and examples. The paragraph should have a strong topic sentence that introduces the main idea and subsequent sentences that elaborate on it, providing specific information or evidence to support the claim. A well-structured paragraph allows for a concise and accurate summary of its key points. 
"""

# Summarize
summary = summarizer(text, max_length=60, min_length=10, do_sample=False)

print(summary[0]['summary_text'])
