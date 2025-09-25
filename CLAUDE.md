# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Chrome browser extension called "Document XLSX Downloader" (version 2.0) that extracts and exports documents from collaboration platforms as XLSX files. It supports Google Sheets and Feishu (飞书).

## Development Setup

**No build system** - this is a vanilla JavaScript Chrome extension that loads directly into Chrome.

### Installation & Testing
- Load extension manually in Chrome's developer mode
- No automated testing framework - test manually by:
  1. Opening extension popup
  2. Navigating to supported platform pages
  3. Triggering CSV download functionality

### Key Files
- `manifest.json` - Chrome extension manifest (Manifest V3)
- `background.js` - Service worker for background operations
- `popup.html` + `popup.js` - Extension UI and logic

## Architecture

### Extension Components
- **Service Worker** (`background.js`): Handles background operations, storage, and download management
- **Popup UI** (`popup.html/js`): User interface for platform selection and download triggering

### Platform Integration
- **Google Sheets**: Direct XLSX export via Google's API endpoints
- **Feishu**: API-based export with access token authentication

### Communication
- Uses Chrome messaging API between components
- Chrome storage API for configuration persistence
- Downloads API for file saving

## Code Conventions

- Mixed English/Chinese comments and UI elements
- Vanilla JavaScript (no frameworks or build tools)
- Chrome Extension Manifest V3 API usage
- Async/await pattern for API calls
- Error handling with user-facing messages

## Permissions
- `activeTab`, `scripting`, `downloads`, `storage`, `tabs`
- Host permissions for Google Sheets and Feishu domains

## Notes
- No automated testing - manual testing required
- No linting/formatting tools configured
- Extension icons are minimal placeholders (1x1 PNG files)
- Bilingual codebase supporting Chinese workplace tools