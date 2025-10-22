# Contributing to Bloomberg Terminal Excel Connector

Thank you for your interest in contributing to the Bloomberg Terminal Excel Connector! This document provides guidelines and instructions for contributing.

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for all contributors.

## How to Contribute

### Reporting Bugs

If you find a bug, please create an issue with:
- Clear description of the bug
- Steps to reproduce
- Expected behavior
- Actual behavior
- Environment details (Node.js version, OS, etc.)

### Suggesting Features

Feature suggestions are welcome! Please create an issue with:
- Clear description of the proposed feature
- Use cases and benefits
- Any implementation ideas you have

### Pull Requests

1. **Fork the repository** and create a new branch for your feature or fix
2. **Install dependencies**: `npm install`
3. **Make your changes** following the coding standards below
4. **Add tests** for new functionality
5. **Run tests**: `npm test`
6. **Run linting**: `npm run lint`
7. **Build the project**: `npm run build`
8. **Commit your changes** with clear, descriptive messages
9. **Push to your fork** and submit a pull request

## Development Setup

```bash
# Clone the repository
git clone https://github.com/prospectfuturesinc/bloomberg-terminal-integrations.git
cd bloomberg-terminal-integrations

# Install dependencies
npm install

# Copy environment configuration
cp config/.env.example config/.env

# Edit .env with your Bloomberg API credentials
# (Required for testing with real Bloomberg API)

# Build the project
npm run build

# Run tests
npm test

# Run examples
npm run dev examples/basic-export.ts
```

## Coding Standards

### TypeScript

- Use TypeScript for all new code
- Enable strict mode
- Provide type definitions for all public APIs
- Avoid using `any` type when possible

### Code Style

- Use 2 spaces for indentation
- Use single quotes for strings
- Add semicolons at end of statements
- Use meaningful variable and function names
- Add JSDoc comments for public APIs

### Example:

```typescript
/**
 * Export market data to Excel file
 * @param securities - Array of securities to export
 * @param fields - Bloomberg fields to retrieve
 * @param fileName - Output file name
 * @returns Promise resolving to file path
 */
async exportMarketData(
  securities: SecurityIdentifier[],
  fields: string[],
  fileName: string
): Promise<string> {
  // Implementation
}
```

### Testing

- Write unit tests for new functionality
- Use Jest for testing
- Aim for high test coverage
- Test both success and error cases

### Commit Messages

Follow conventional commit format:

```
type(scope): subject

body

footer
```

Types:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `test`: Adding or updating tests
- `chore`: Maintenance tasks

Example:
```
feat(excel): add conditional formatting support

Added support for color scales and data bars in Excel exports.
Includes new formatting options in ExcelFormattingOptions type.

Closes #123
```

## Project Structure

```
bloomberg-terminal-integrations/
├── src/                  # Source code
│   ├── connectors/       # Bloomberg and Excel connectors
│   ├── types/            # TypeScript type definitions
│   ├── utils/            # Utility functions
│   └── index.ts          # Main entry point
├── examples/             # Example scripts
├── tests/                # Test files
├── config/               # Configuration files
└── docs/                 # Documentation
```

## Adding New Features

When adding a new feature:

1. **Plan the API**: Design clear, intuitive interfaces
2. **Add types**: Create TypeScript types in appropriate files
3. **Implement**: Write the implementation
4. **Test**: Add comprehensive tests
5. **Document**: Update README and add JSDoc comments
6. **Examples**: Add example usage in `examples/`

## Testing with Bloomberg API

To test with real Bloomberg API:

1. Obtain Bloomberg Terminal API credentials
2. Configure `.env` file with your credentials
3. Run examples: `npm run dev examples/basic-export.ts`

Note: Some tests may be skipped without real Bloomberg API access.

## Documentation

- Keep README.md up to date
- Add JSDoc comments for public APIs
- Update CHANGELOG.md for significant changes
- Add examples for new features

## Questions?

If you have questions about contributing, feel free to:
- Create a GitHub issue
- Email: support@prospectfutures.com

Thank you for contributing!
