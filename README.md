# Shinsen Bot - LINE Excel Screenshot Detector

A LINE bot server that detects if a message contains an image and determines whether the image is an Excel screenshot using Claude AI vision capabilities.

## Features

- Detects image messages from LINE
- Uses Claude AI vision to identify Excel/spreadsheet screenshots
- Responds to users with detection results

## Prerequisites

- Node.js (v16 or higher)
- LINE Messaging API channel
- Anthropic API key

## Setup

1. Install dependencies:
```bash
npm install
```

2. Create a `.env` file based on `.env.example`:
```bash
cp .env.example .env
```

3. Configure your environment variables in `.env`:
   - `LINE_CHANNEL_ACCESS_TOKEN`: Get this from LINE Developers Console
   - `LINE_CHANNEL_SECRET`: Get this from LINE Developers Console
   - `ANTHROPIC_API_KEY`: Get this from Anthropic Console
   - `PORT`: Server port (default: 3000)

## Getting LINE Credentials

1. Go to [LINE Developers Console](https://developers.line.biz/console/)
2. Create a new provider or select existing one
3. Create a new Messaging API channel
4. In the channel settings:
   - Get your Channel Secret
   - Issue a Channel Access Token (long-lived)
   - Set webhook URL to `https://your-domain.com/webhook`
   - Enable "Use webhook"

## Getting Anthropic API Key

1. Go to [Anthropic Console](https://console.anthropic.com/)
2. Sign up or log in
3. Navigate to API Keys section
4. Create a new API key

## Running the Server

Development mode with auto-reload:
```bash
npm run dev
```

Production mode:
```bash
npm start
```

The server will start on the configured PORT (default: 3000).

## Webhook Setup

1. Make your server publicly accessible (use ngrok for local development):
```bash
ngrok http 3000
```

2. Copy the HTTPS URL from ngrok
3. Set the webhook URL in LINE Developers Console:
   - Format: `https://your-ngrok-url.ngrok.io/webhook`

## How It Works

1. User sends an image to the LINE bot
2. Server receives the webhook event at `/webhook`
3. Server checks if the message type is `image`
4. If it's an image:
   - Downloads the image content from LINE
   - Sends the image to Claude AI for analysis
   - Claude determines if it's an Excel/spreadsheet screenshot
   - Bot replies with the detection result
5. If it's not an image, the message is ignored

## API Endpoints

- `POST /webhook` - LINE webhook endpoint
- `GET /health` - Health check endpoint

## Project Structure

```
shinsen_bot/
├── server.js           # Main server file
├── package.json        # Dependencies and scripts
├── .env.example        # Example environment variables
├── .gitignore         # Git ignore file
└── README.md          # This file
```

## Key Functions

- `handleEvent()` - Processes incoming LINE events in server.js:17
- `getImageContent()` - Downloads image from LINE servers in server.js:54
- `detectExcelScreenshot()` - Uses Claude AI to detect Excel screenshots in server.js:68

## Testing

Send an image to your LINE bot:
- Excel screenshot: Bot responds "This is an Excel screenshot!"
- Other image: Bot responds "This is an image, but not an Excel screenshot."

## Troubleshooting

- Ensure all environment variables are set correctly
- Check that your webhook URL is HTTPS and accessible
- Verify LINE webhook is enabled in LINE Developers Console
- Check server logs for error messages

## License

MIT
