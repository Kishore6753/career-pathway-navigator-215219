# Lightweight React Template for KAVIA

This project provides a minimal React template with a clean, modern UI and minimal dependencies.

## Features

- **Lightweight**: No heavy UI frameworks - uses only vanilla CSS and React
- **Modern UI**: Clean, responsive design with KAVIA brand styling
- **Fast**: Minimal dependencies for quick loading times
- **Simple**: Easy to understand and modify

## Getting Started

In the project directory, you can run:

### `npm start`

Runs the app in development mode.\
Open [http://localhost:3000](http://localhost:3000) to view it in your browser.

### `npm test`

Launches the test runner in interactive watch mode.

### `npm run build`

Builds the app for production to the `build` folder.\
It correctly bundles React in production mode and optimizes the build for the best performance.

## Python-based Supabase Seeding (No Node.js Required)

A Python CLI seeding utility is included to upsert data into Supabase using the provided Excel attachments via REST.

What it seeds (selectively):
- roles
- competencies
- role_competencies
- role_descriptions
- role_adjacency
- learning_resources

Environment variables (must be set in your shell; do not read .env directly):
- SUPABASE_URL
- SUPABASE_SERVICE_ROLE_KEY

Installation:
- Python 3.10+ recommended
- Install dependencies:
  pip install -r career-pathway-navigator-215219/requirements.txt

CLI usage:
- Show help:
  python3 career-pathway-navigator-215219/career_navigator_frontend/scripts/seed.py --help

Examples:
- Seed everything (default sources from attachments/):
  SUPABASE_URL="https://<project>.supabase.co" \
  SUPABASE_SERVICE_ROLE_KEY="<service_role>" \
  python3 career-pathway-navigator-215219/career_navigator_frontend/scripts/seed.py

- Seed a subset (roles + competencies + role_competencies only):
  SUPABASE_URL="https://<project>.supabase.co" \
  SUPABASE_SERVICE_ROLE_KEY="<service_role>" \
  python3 career-pathway-navigator-215219/career_navigator_frontend/scripts/seed.py \
    --only roles,competencies,role_competencies

- Override source files:
  python3 career-pathway-navigator-215219/career_navigator_frontend/scripts/seed.py \
    --competency "/path/to/Competency_mapping.xlsx" \
    --adjacency "/path/to/CA_Role_Adjacency.xlsx" \
    --roles "/path/to/Role_Navigator_Worksheet.xlsx" \
    --resources "/path/to/Learning_Resources.xlsx" \
    --cardsGlob "/path/to/Role_Card_*"

Default input paths (already present in this repo’s attachments):
- Competencies: /home/kavia/workspace/code-generation/attachments/20251127_090403_Competency_mapping.xlsx
- Role adjacency (preferred): /home/kavia/workspace/code-generation/attachments/20251127_090402_CA_Role_Adjacency29.xlsx
- Role adjacency (alternate): /home/kavia/workspace/code-generation/attachments/20251127_090401_CA_Role_Adjacency.xlsx
- Role navigator worksheet: /home/kavia/workspace/code-generation/attachments/20251127_090406_Role_Navigator_Worksheet.xlsx
- Role cards (text): /home/kavia/workspace/code-generation/attachments/2025*Role_Card_*

Mapping assumptions:
- Natural keys (codes) are slugs of human-readable names when explicit keys are absent:
  - roles: code = slug(title or column header)
  - competencies: code = slug(Competency Code or Competency Name)
- role_competencies levels map flexibly:
  B/Basic/Beginner/Foundation=1, P/Proficient/Intermediate=2, A/Advanced=3, M/Master/Expert/Authority=4
  Numeric cells are coerced into 1..4
- role_adjacency:
  - Row-wise sheets supported (Source/Target/Score)
  - Matrix sheets supported (first row = targets, first col = sources, cells = scores)
- role_descriptions:
  - Populated from Role_Card_*.txt
  - roles.description is set to a short summary; full text is inserted into role_descriptions with source='card'
- learning_resources:
  - Expects Title and URL columns
  - Optionally maps to a competency via “Competency Code” or “Competency” column

Idempotency and upsert keys:
- roles: (code)
- competencies: (code)
- role_competencies: (role_id, competency_id)
- role_adjacency: (source_role_id, target_role_id)
- role_descriptions: (role_id, source)
- learning_resources: (competency_id, url)

Retries:
- REST calls use exponential backoff and log transient errors before retrying.

Schema alignment:
- Matches the included SQL at supabase/schema.sql, including role_descriptions and constraints.

## Customization

### Colors

The main brand colors are defined as CSS variables in `src/App.css`:

```css
:root {
  --kavia-orange: #E87A41;
  --kavia-dark: #1A1A1A;
  --text-color: #ffffff;
  --text-secondary: rgba(255, 255, 255, 0.7);
  --border-color: rgba(255, 255, 255, 0.1);
}
```

### Components

This template uses pure HTML/CSS components instead of a UI framework. You can find component styles in `src/App.css`. 

Common components include:
- Buttons (`.btn`, `.btn-large`)
- Container (`.container`)
- Navigation (`.navbar`)
- Typography (`.title`, `.subtitle`, `.description`)

## Learn More

To learn React, check out the [React documentation](https://reactjs.org/).

### Code Splitting

This section has moved here: [https://facebook.github.io/create-react-app/docs/code-splitting](https://facebook.github.io/create-react-app/docs/code-splitting)

### Analyzing the Bundle Size

This section has moved here: [https://facebook.github.io/create-react-app/docs/analyzing-the-bundle-size](https://facebook.github.io/create-react-app/docs/analyzing-the-bundle-size)

### Making a Progressive Web App

This section has moved here: [https://facebook.github.io/create-react-app/docs/making-a-progressive-web-app](https://facebook.github.io/create-react-app/docs/making-a-progressive-web-app)

### Advanced Configuration

This section has moved here: [https://facebook.github.io/create-react-app/docs/advanced-configuration](https://facebook.github.io/create-react-app/docs/advanced-configuration)

### Deployment

This section has moved here: [https://facebook.github.io/create-react-app/docs/deployment](https://facebook.github.io/create-react-app/docs/deployment)

### `npm run build` fails to minify

This section has moved here: [https://facebook.github.io/create-react-app/docs/troubleshooting#npm-run-build-fails-to-minify](https://facebook.github.io/create-react-app/docs/troubleshooting#npm-run-build-fails-to-minify)
