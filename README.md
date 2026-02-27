# Mbox JSONL Reviewer (Very Simple)

Tiny web app that:
1. Uploads an `.mbox` file
2. Converts messages to JSONL
3. Sends records to an LLM for a quick review

## Run

```bash
npm install
npm start
```

Open: `http://localhost:3000`

## Optional: live LLM

Set your API key before running:

```bash
export OPENAI_API_KEY="your_key_here"
```

If no key is set, the app returns a mock review so the UI still works.

## Notes

- Parser is intentionally basic to keep the app simple.
- Large mbox files can be slow; this is a starter UI, not production-grade parsing.
