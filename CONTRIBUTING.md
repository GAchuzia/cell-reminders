# Contributing to Cell Reminders

## Setup

1. Fork the repository
2. Clone your fork: `git clone https://github.com/yourusername/cell-reminders.git`
3. Install dependencies: `npm install`
4. Create feature branch: `git checkout -b feature/your-feature`

## Submitting Changes

1. Commit changes with [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/) format:

   ```bash
   git commit -m "feat: add new reminder feature"
   git commit -m "fix: resolve calendar sync issue"
   git commit -m "docs: update installation guide"
   ```

2. Push to your fork: `git push origin feature/your-feature`

3. Create pull request using the template

## Branch Strategy

- **main**: Production-ready code
- **feature/***: New features
- **bugfix/***: Bug fixes
- **hotfix/***: Critical fixes

## Commit Message Format

Use conventional commits:

``` git
type(scope): description

[optional body]

[optional footer]
```

Examples:

- `feat(ui): add loading indicators`
- `fix(calendar): resolve timezone issues`
- `docs(readme): update installation steps`
- `refactor(utils): improve error handling`
