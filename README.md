# Data Science & Public Policy Analytics Portfolio

Professional portfolio website showcasing data science and analytics projects developed for public education management systems.

## Overview

This portfolio presents projects built for state and municipal education secretariats in Brazil, including:

- Educational risk indicator platforms
- Student monitoring dashboards
- Assessment management systems
- Data integration pipelines
- Learning recovery tracking tools
- School prioritization systems

## Tech Stack

- HTML5, CSS3 (Flexbox/Grid), JavaScript
- Google Fonts (Inter, Open Sans)
- Font Awesome icons
- No frameworks — pure HTML/CSS/JS for maximum portability

## Deploy on GitHub Pages

1. Create a new GitHub repository (e.g., `portfolio`)
2. Push this project to the repository:

```bash
cd portfolio-site
git init
git add .
git commit -m "Initial portfolio site"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/portfolio.git
git push -u origin main
```

3. Go to **Settings > Pages** in your GitHub repository
4. Under **Source**, select `main` branch and `/ (root)` folder
5. Click **Save** — your site will be live at `https://YOUR_USERNAME.github.io/portfolio/`

## Adding New Projects

1. Create a new file in the `projects/` folder (e.g., `project9.html`)
2. Copy the structure from any existing project page
3. Update the content: title, context, solution, data sources, technologies, and impact
4. Add a new card in `projects.html` linking to the new page
5. Optionally add it to the featured section on `index.html`

## Customization

- **Colors:** Edit CSS variables in `assets/css/styles.css` (`:root` block)
- **Profile photo:** Replace the placeholder in `about.html` with `<img src="assets/images/profile.jpg">`
- **Links:** Update LinkedIn, GitHub, and email links across all pages
- **CV:** Add a PDF to `assets/` and update the download link in `about.html`

## Structure

```
├── index.html          # Home / landing page
├── about.html          # Professional profile
├── skills.html         # Technical competencies
├── projects.html       # Project gallery
├── contact.html        # Contact information & form
├── assets/
│   ├── css/styles.css  # Main stylesheet
│   ├── js/script.js    # Navigation & animations
│   └── images/         # Profile photo, screenshots
├── projects/
│   ├── project1.html   # Educational Risk Indicators Platform
│   ├── project2.html   # Student Migration Monitoring Dashboard
│   ├── project3.html   # Education Performance Dashboard
│   ├── project4.html   # School Assessment Management System
│   ├── project5.html   # Learning Recovery Monitoring Dashboard
│   ├── project6.html   # Temporal Risk Monitoring System
│   ├── project7.html   # School Prioritization System for SAEB
│   └── project8.html   # Educational Indicators Panel (IDEB & Census)
└── README.md
```

## License

This portfolio is for professional use. Project descriptions reflect real work developed for public education institutions.
