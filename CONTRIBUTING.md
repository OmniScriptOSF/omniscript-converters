# Contributing to OmniScript Converters

Thank you for your interest in contributing to **OmniScript Format (OSF) Converters**! This package provides format conversion from OSF to PDF, DOCX, PPTX, and XLSX.

---

## 🚀 Getting Started

### 1️⃣ Fork the repository

Click the **Fork** button at the top right of the [omniscript-converters](https://github.com/OmniScriptOSF/omniscript-converters) repository page.

### 2️⃣ Clone your fork locally

```bash
git clone https://github.com/your-username/omniscript-converters.git
cd omniscript-converters
git checkout -b my-feature-branch
```

### 3️⃣ Install dependencies

```bash
# Install pnpm if you haven't already
npm install -g pnpm

# Install project dependencies
pnpm install

# Build the package
pnpm run build
```

### 4️⃣ Make your changes

- Follow the coding style enforced by ESLint and Prettier
- Add/update tests where relevant
- Update documentation if needed
- Ensure type safety with TypeScript

### 5️⃣ Run quality checks

```bash
# Run all tests
pnpm test

# Type checking
pnpm run typecheck

# Lint your code
pnpm run lint

# Format your code
pnpm run format
```

### 6️⃣ Commit and push

```bash
git add .
git commit -m "feat: describe your change concisely"
git push origin my-feature-branch
```

### 7️⃣ Open a Pull Request

Go to your fork on GitHub and click **Compare & pull request**.

---

## 💡 Contribution Types

✅ Fix bugs in converters  
✅ Add support for new output formats  
✅ Improve conversion quality  
✅ Add tests for edge cases  
✅ Improve documentation  
✅ Optimize performance  

---

## ✨ Guidelines

### Commit Message Convention

We follow [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` - A new feature (e.g., new format support)
- `fix:` - A bug fix
- `docs:` - Documentation only changes
- `test:` - Adding or updating tests
- `perf:` - Performance improvements
- `refactor:` - Code refactoring

**Examples:**
```
feat: add table support to PDF converter
fix: handle empty cells in XLSX export
docs: update converter API examples
```

### Testing Requirements

- All new features must include tests
- Test all supported formats (PDF, DOCX, PPTX, XLSX)
- Test edge cases and error handling
- Aim for 80%+ code coverage

### Pull Request Process

1. Target the `main` branch
2. Ensure all CI checks pass
3. Request review from maintainers
4. Address any review feedback

### All contributors must follow our [Code of Conduct](CODE_OF_CONDUCT.md)

---

## 🤝 Community

Join our discussions on [GitHub Discussions](https://github.com/OmniScriptOSF/omniscript-core/discussions).

---

## 📚 Key Technologies

- **Language**: TypeScript 5.x
- **PDF**: pdfkit
- **DOCX**: docx
- **PPTX**: pptxgenjs
- **XLSX**: exceljs
- **Testing**: Vitest

---

## 📄 License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
