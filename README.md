# EventNest — Event Management Software

Run the show. Not six WhatsApp groups. AI-powered checklists, vendor approvals, and team coordination for event agencies and wedding planners.

## Features

Inspired by eventops.co.in, this system includes:

- **Proposal Generator**: Build proposals your client can actually picture — moodboard, creative direction, and cost breakdown in minutes
- **Sub-Event Management**: Handle wedding functions (Mehendi, Sangeet, Wedding, Reception) as separate sub-events with their own checklists
- **Specialized Vendor Categories**: Real event production categories (Flex Printing, Digital Printing, Truss, Fabrication, Genset, Power, etc.)
- **AI Creative Studio**: Generate stage designs, logo designs, moodboards, color palettes, and lighting designs
- **Team Management**: Add and manage team members with specific roles
- **Real Venue Data**: Track venues with confidence scores
- **Runsheet Management**: Detailed event-day schedules
- **Budget Tracking**: Category-specific budget items with production-focused categories
- **Checklist Engine**: Task management with team member assignment

## Tech Stack

- **Frontend & Backend**: Streamlit (all-in-one)
- **Data Storage**: Excel (openpyxl, pandas)
- **AI/LLM**: Hugging Face API (free, using Mistral-7B-Instruct)

## Project Structure

```
eventops/
├── app.py                   # Main Streamlit application
├── data/
│   └── eventops_data.xlsx   # Excel database
├── requirements.txt
└── README.md
```

## Local Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

3. The app will open in your browser at `http://localhost:8501`

## AI Features

To enable AI features, you'll need a Hugging Face API key (free):
1. Go to [huggingface.co/settings/tokens](https://huggingface.co/settings/tokens)
2. Create a new token
3. Enter your API key in the sidebar under "Hugging Face API Key"

The AI uses the Mistral-7B-Instruct-v0.2 model for:
- Proposal generation
- Creative direction (stage design, logo, moodboard)
- Color palette suggestions
- Lighting design recommendations

## Deployment

### Streamlit Cloud (Free)

1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click "New app"
4. Connect your GitHub repository
5. Set main file to `app.py`
6. Add your Hugging Face API key in secrets
7. Click "Deploy"

Note: Streamlit Cloud doesn't provide persistent file storage. Data will reset on redeploy. For production persistence, consider cloud storage or database integration.
