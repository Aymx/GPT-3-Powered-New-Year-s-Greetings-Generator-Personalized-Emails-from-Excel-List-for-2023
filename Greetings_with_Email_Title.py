# "This Python code reads data from an Excel file, including names of business partners and friends, 
# things to mention, language, and tone. It then uses this information to generate personalized New Year's 
# greetings emails for 2023, with a customizable language and tone. The code also creates a title for the email.
# It uses the OpenAI text-davinci-003 engine
# It reads the data from the attached file list.xlsx and writes the output to list_reply.xlsx



import openai
import pandas as pd

# Replace YOUR_API_KEY with your actual API key for the GPT-3 API
openai.api_key = "YOUR_API_KEY"

# Use GPT-3 to generate a thank you message
def generate_email_message(prompt):
    # Add the following line
    model_engine = "text-davinci-003"
    prompt = (f"{prompt}"
             )

    completions = openai.Completion.create(
        engine=model_engine,
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.7,
    )

    message = completions.choices[0].text
    return message

# Read the Excel file
df = pd.read_excel("list.xlsx")


# Add 'email_message' column to the dataframe
df['email_message'] = None

# Iterate through each row in the dataframe
for index, row in df.iterrows():
    name = row["name"]
    things_to_mention = row["things_to_mention"]
    language = row["language"]
    tone = row["tone"]
        
    # Generate the thank you message
    prompt = (f"Write a greeting Email for 2023 to {name}. Don't forget to mention {things_to_mention} in the message. The language of the message should be written in {language}. The tone of the message should be {tone}. At the end of the Email, sign it with my name (Max Mustermann)."
             )
    message = generate_email_message(prompt)

    # Update the email_message column with the generated message
    df.at[index, "email_message"] = message

    # Generate the Title for the Email
    email_message = row["email_message"]
    prompt = (f"Write a title for an Email like (Happy New Year 2023) or something similar and you can mention things from the following text, the language of the title should be written in {language}: {email_message}"
             )
    message = generate_email_message(prompt)

    # Update the email_message column with the generated message
    df.at[index, "email_title"] = message

# Save the updated dataframe to the Excel file
df.to_excel("list_reply.xlsx", index=False)