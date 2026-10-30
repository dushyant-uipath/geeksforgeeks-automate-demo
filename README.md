<p align="center">
  <img src="assets/uipath-logo.jpeg" alt="UiPath" width="160" />
</p>

<h1 align="center">DevPulse</h1>

<p align="center">
  A Maestro Automate demo for the GeeksforGeeks webinar: an agent writes a developer newsletter, a human reviews it, and Outlook sends it.
</p>

## What it does

DevPulse pulls the last 48 hours of Hacker News stories on a topic, asks an agent to pick the five most useful ones and write the issue, then pauses for a reviewer in Action Center. On **Send** the issue goes out through the Outlook 365 connector. On **Discard** the run ends.

| Step | Node | What it does |
|---|---|---|
| 1 | `start` | Manual trigger with two inputs: `recipient` and `topic` |
| 2 | `fetchNews` | HTTP GET to the Hacker News Algolia search API |
| 3 | `prep` | Script that maps hits to stories and builds `storiesJson` |
| 4 | `editor` | Inline agent (`gpt-5.4`) that returns `subject`, `digest` and `bodyHtml` |
| 5 | `review` | Human task in the `devpulse-review` coded action app, outcomes `Send` and `Discard` |
| 6 | `verdict` | Switch on the review outcome |
| 7 | `sendEmail` | Outlook 365 Send Email, `Send` branch only |
| 8 | End | Returns `status` (`sent` or `discarded`) and `subject` |

## Repository layout

```
DevPulse/
  DevPulse.uipx              solution manifest
  DevPulse/                  Maestro Automate flow project
    .flow-sdk/DevPulse.flow.ts   flow source (TypeScript builder SDK)
    DevPulse.flow                compiled flow
  devpulse-review/           coded action app (React + Vite) used by the review step
  resources/                 solution resources
Maestro-Automate-GFG-Webinar.pdf   webinar deck
```

## Run it

1. Install the UiPath CLI and sign in with `uip login`.
2. Create an Outlook 365 connection in Integration Service. The Send branch needs it.
3. Build the review app: `npm install` then `npm run build` inside `DevPulse/devpulse-review`.
4. Upload the solution to Studio Web: `uip solution upload DevPulse`.
5. Debug the flow with this input:

```json
{"recipient": "you@example.com", "topic": "AI agents"}
```

An empty `topic` falls back to `AI agents`. Other topics to try: `LLM evals`, `Rust`, `Kubernetes`.

## Edit with a coding agent

The flow source is `DevPulse/DevPulse/.flow-sdk/DevPulse.flow.ts`. Open the `DevPulse` folder in VS Code with the UiPath coding agent skills installed, ask for a change in plain English, then recompile:

```
uip maestro flow compile .flow-sdk/DevPulse.flow.ts -o DevPulse.flow
uip maestro flow validate DevPulse.flow
```

## Data source

Stories come from Hacker News through the [Algolia Search API](https://hn.algolia.com/api). It needs no API key.
