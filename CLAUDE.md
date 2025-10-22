# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Microsoft Word Add-in built with React, TypeScript, and the Office.js API. It provides two main features:
- **Document Explorer**: Tools to create demo documents and analyze Word document content
- **Variable Manager**: Extract, search, highlight, and replace template variables in documents

## Build & Development Commands

### Core Commands
```bash
npm run dev-server      # Start development server with hot reload (https://localhost:3000)
npm run build           # Production build
npm run build:dev       # Development build
npm run watch           # Development build with file watching
```

### Office Add-in Commands
```bash
npm start               # Start debugging the add-in in Word (uses manifest.xml)
npm stop                # Stop debugging session
npm run validate        # Validate manifest.xml
```

### Code Quality
```bash
npm run lint            # Check for linting issues
npm run lint:fix        # Auto-fix linting issues
npm run prettier        # Format code with Prettier
```

### Authentication (for M365 development)
```bash
npm run signin          # Sign in to M365 account
npm run signout         # Sign out from M365 account
```

## Architecture

### Feature-Based Structure

The codebase follows a **feature-based architecture** where each major functionality is isolated in its own module under `src/taskpane/features/`:

```
src/taskpane/features/
├── document-explorer/
│   ├── hooks/use-document-explorer.ts      # Office.js API logic
│   ├── components/                         # Feature-specific UI components
│   ├── document-explorer-tab.tsx           # Main tab component
│   └── index.ts                            # Public exports
└── variable-manager/
    ├── types/index.ts                      # TypeScript interfaces
    ├── hooks/use-variable-manager.ts       # Business logic & Word API
    ├── components/                         # Feature-specific UI components
    ├── variable-manager-tab.tsx            # Main tab component
    └── index.ts                            # Public exports
```

**Key principles:**
- Each feature is self-contained with its own hooks, types, and components
- Business logic lives in custom hooks (e.g., `use-variable-manager.ts`)
- All Word API calls (`Word.run()`) are isolated in hooks, never in UI components
- Features export only what's needed via `index.ts`

### File Naming Convention

All files use **kebab-case**: `use-variable-manager.ts`, `pattern-config-card.tsx`, etc.

### Main Application

`src/taskpane/components/App.tsx` is the orchestrator:
- Imports feature tabs from `features/`
- Manages global state (status messages, active tab)
- Provides layout and navigation

### UI Components

Reusable UI components are in `src/components/ui/` (Radix UI + Tailwind CSS):
- These are generic, shadcn/ui-style components
- Feature-specific components live within their feature directory

### Path Aliases

Webpack and TypeScript are configured with `@/` as an alias for `src/`:
```typescript
import { Button } from "@/components/ui/button";
```

## Office.js Integration

### Key Patterns

1. **Always wrap Office.js calls in `Word.run()`:**
   ```typescript
   await Word.run(async (context) => {
     const body = context.document.body;
     body.load("text");
     await context.sync();
     // Work with loaded properties
   });
   ```

2. **Load properties before accessing them:**
   ```typescript
   selection.load("text,font,style");
   await context.sync();
   console.log(selection.text); // Now accessible
   ```

3. **Batch operations for performance:**
   ```typescript
   // Load all items at once, sync once
   const paragraphs = context.document.body.paragraphs;
   paragraphs.load("items,text,style");
   await context.sync();
   ```

### Feature Hook Pattern

When adding new Word API functionality:
1. Create/update a hook in `features/[feature-name]/hooks/`
2. Accept an `onStatusChange` callback for user feedback
3. Implement async functions that use `Word.run()`
4. Return functions to be called by UI components

Example:
```typescript
export const useMyFeature = (onStatusChange: (status: string) => void) => {
  const doSomething = async () => {
    try {
      await Word.run(async (context) => {
        // Office.js logic here
      });
      onStatusChange("✅ Success");
    } catch (error) {
      console.error(error);
      onStatusChange("❌ Error");
    }
  };

  return { doSomething };
};
```

## Adding New Features

To add a new feature:

1. Create feature directory structure:
   ```
   src/taskpane/features/my-feature/
   ├── index.ts
   ├── my-feature-tab.tsx
   ├── hooks/use-my-feature.ts
   ├── components/
   └── types/ (if needed)
   ```

2. Implement the hook with Word API logic
3. Create UI components that consume the hook
4. Assemble components in the main tab file
5. Export the tab from `index.ts`
6. Import and add tab to `App.tsx`

## Development Notes

### HTTPS Development Server

The dev server runs on HTTPS with self-signed certificates (required for Office Add-ins):
- Certificates are auto-generated by `office-addin-dev-certs`
- Dev server runs on port 3000 by default
- Accept the certificate warning in your browser on first run

### Manifest File

`manifest.xml` defines the add-in configuration:
- Points to `https://localhost:3000/` in development
- Update `urlProd` in `webpack.config.js` for production deployment
- Validate changes with `npm run validate`

### Styling

- Tailwind CSS 4.x with PostCSS
- Component variants via `class-variance-authority`
- Global styles in `src/taskpane/styles/globals.css`

### TypeScript Configuration

- Target: ES5 (for IE11 compatibility per browserslist)
- Module: ES2020
- JSX: React
- Strict type checking enabled
