from openai import OpenAI


class OpenAIClient:
    def __init__(self, api_key, model):
        self.api_key = api_key
        self.model = model

    def generate(self, prompt) -> str:
        client = OpenAI(api_key=self.api_key)

        completion = client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are a professional presentation creator assistant."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=2000
        )

        return completion.choices[0].message.content
