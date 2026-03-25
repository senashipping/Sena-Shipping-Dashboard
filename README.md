# SENA Ship Dashboard

Full-stack ship management and compliance dashboard for SENA Shipping.

This project includes:
- A Node.js + Express backend API with MongoDB.
- A React + TypeScript frontend with role-based navigation.
- Dynamic form building and submission workflows.
- Notification and cron-based compliance checks.

## Table of Contents

1. Project Overview
2. Core Features
3. Architecture
4. Project Structure
5. Tech Stack
6. Prerequisites
7. Local Setup
8. Environment Variables
9. Available Scripts
10. Roles and Access Control
11. Functional Modules
12. API Summary
13. Seed Data and Default Credentials
14. Deployment Notes
15. Additional Documentation
16. Known Notes and Gotchas

## Project Overview

SENA Ship Dashboard is designed to manage maritime forms, assignments, submissions, review workflows, and related notifications across different user roles.

The system supports:
- Admin and super admin operations for forms, ships, users, and submission reviews.
- User workflows for completing and tracking assigned forms.
- Category-based form distribution, including a shared deck+engine category.
- Automated reminder/expiration notifications through scheduled tasks.

## Core Features

- JWT authentication with access and refresh tokens.
- Role-based authorization (`user`, `admin`, `super_admin`).
- Dynamic form builder with multiple field types.
- Signature upload field with image validation (stored as base64).
- Submission lifecycle (draft, submitted, approved, rejected, expired).
- Category-based form filtering and assignment logic.
- In-app notifications and notification management.
- Daily and weekly cron jobs for compliance and cleanup operations.

## Architecture

### Backend

- Entry point: `Backend/index.js`
- Framework: Express
- Database: MongoDB via Mongoose
- Route modules:
  - `/api/auth`
  - `/api/admin`
  - `/api/forms`
  - `/api/submissions`
  - `/api/notifications`
- Middleware:
  - Auth/JWT verification
  - Role authorization
  - Joi request validation
  - Centralized error handling
- Scheduled services:
  - Expiration and status checks (daily)
  - Old notification cleanup (weekly)
  - Startup missed-notification recovery

### Frontend

- Entry point: `FrontEnd/src/main.tsx`
- App routing: `FrontEnd/src/App.tsx`
- Framework: React + TypeScript + Vite
- State and providers:
  - Auth context
  - Notification context
  - Theme context
- UI system:
  - Tailwind CSS
  - Radix UI primitives
- API integration:
  - Axios service with auth header injection
  - Auto refresh-token flow on `401`

## Project Structure

```text
Sena/
  Backend/
    controllers/
    middleware/
    models/
    routes/
    services/
    utils/
    scripts/
    index.js
  FrontEnd/
    src/
      api/
      components/
      contexts/
      layouts/
      pages/
      types/
    public/
    vite.config.ts
    vercel.json
```

## Tech Stack

### Backend

- Node.js
- Express 5
- MongoDB + Mongoose
- JWT (`jsonwebtoken`)
- Password hashing (`bcryptjs`)
- Validation (`joi`)
- Security/logging (`helmet`, `cors`, `morgan`)
- Scheduling (`node-cron`)

### Frontend

- React 18
- TypeScript
- Vite
- React Router v6
- Axios
- React Query
- Tailwind CSS
- Radix UI
- DnD Kit (form-builder interactions)

## Prerequisites

- Node.js 18+
- npm (backend)
- pnpm (frontend)
- MongoDB instance (local or hosted)

## Local Setup

### 1. Clone and install

```bash
git clone <your-repository-url>
cd Sena
```

Backend dependencies:

```bash
cd Backend
npm install
```

Frontend dependencies:

```bash
cd ../FrontEnd
pnpm install
```

### 2. Configure environment files

- Copy `Backend/.env.example` to `Backend/.env`
- Copy `FrontEnd/.env.example` to `FrontEnd/.env`

### 3. Start backend

```bash
cd Backend
npm run dev
```

### 4. Start frontend

```bash
cd FrontEnd
pnpm dev
```

Default local URLs:
- Frontend: `http://localhost:3000`
- Backend API: `http://localhost:8080`
- Health check: `http://localhost:8080/api/health`

## Environment Variables

### Backend (`Backend/.env`)

Required:
- `MONGO_URI`
- `JWT_SECRET`
- `JWT_REFRESH_SECRET`

Common:
- `JWT_EXPIRE` (default example: `7d`)
- `JWT_REFRESH_EXPIRE` (default example: `30d`)
- `PORT` (example file uses `5000`; runtime fallback is `8080`)
- `NODE_ENV`
- `FRONTEND_URL`
- `TIMEZONE`

Optional:
- `RATE_LIMIT_WINDOW_MS`
- `RATE_LIMIT_MAX_REQUESTS`
- `MAX_FILE_SIZE`
- `UPLOAD_PATH`

### Frontend (`FrontEnd/.env`)

- `VITE_API_URL` (example: `http://localhost:8080/api`)

## Available Scripts

### Backend (`Backend/package.json`)

- `npm start` -> run server with Node
- `npm run dev` -> run server with nodemon

### Frontend (`FrontEnd/package.json`)

- `pnpm dev` -> start Vite dev server
- `pnpm build` -> TypeScript check + Vite build
- `pnpm build:prod` -> production-mode build
- `pnpm preview` -> preview built app
- `pnpm lint` -> run ESLint

## Roles and Access Control

- `super_admin`
  - Highest role
  - Can perform super-admin-protected actions (for example destructive admin operations)

- `admin`
  - Access to admin dashboard modules
  - Can manage forms/categories and review submissions

- `user`
  - Access to user dashboard
  - Can fill and submit assigned forms

Additional segmentation:
- `userType`: `deck` or `engine`
- Form category targeting uses role + userType rules

## Functional Modules

### Auth

- Register, login, profile, change password, refresh token, logout.

### Admin Console

- Dashboard analytics.
- User and ship management.
- Pending submission review and status stats.

### Form Management

- Category CRUD (admin/super-admin constraints on specific actions).
- Form CRUD, duplicate, active toggle.
- Category-specific listing and user-status listing.

### Form Builder

- Drag-and-drop field management.
- Section support for mixed form layouts.
- Field types include:
  - `text`, `email`, `number`, `date`, `datetime-local`, `time`
  - `textarea`, `select`, `checkbox`, `radio`
  - `phone`, `url`, `signature`

### Signature Upload

- Image-only upload (`jpeg`, `jpg`, `png`, `gif`, `webp`).
- Client-side max size: 2MB.
- Data stored as base64 string in submission payload.

### Submissions

- Draft/save/update flow for users.
- Admin review endpoints for approve/reject.
- Status summary and user dashboard data.

### Notifications

- User notifications with read/read-all/delete.
- Admin-triggerable system notifications.
- Debug and cleanup utility routes for expiration logic.

### Scheduled Jobs

- Daily at 12:00 UTC:
  - Expiration notifications
  - Form status notifications
  - Unfilled-form checks
  - Duplicate cleanup routine
- Weekly Sunday 02:00 (configured timezone):
  - Old notification cleanup

## API Summary

Base URL: `/api`

- Auth: `/auth`
  - login/register/profile/password/refresh/logout
- Admin: `/admin`
  - dashboard, users, ships, submission stats/pending
- Forms: `/forms`
  - categories, forms, user-status, duplicate, toggle-status, migrate-validity-period
- Submissions: `/submissions`
  - CRUD, dashboard, status-summary, review
- Notifications: `/notifications`
  - list, stats, read, read-all, delete, system, trigger/debug maintenance routes

For detailed payload examples, see the backend route/controller code and the docs listed below.

## Seed Data and Default Credentials

In non-production mode, backend startup seeds default records if missing.

Seeded defaults include:
- Super admin and admin users
- Categories (`eng`, `deck`, `mlc`, `isps`, `drill`, `deck_engine`)
- Demo ships and deck/engine user accounts

Default credentials (development seeding):
- Super Admin: `superadmin@senashipping.com` / `superadmin123`
- Admin: `admin@senashipping.com` / `admin123`
- Deck User 1: `deck@senashipping.com` / `deck123`
- Engine User 1: `engine@senashipping.com` / `engine123`
- Deck User 2: `deck2@senashipping.com` / `deck123`
- Engine User 2: `engine2@senashipping.com` / `engine123`

## Deployment Notes

- Frontend deployment config is in `FrontEnd/vercel.json`.
- Current file includes a hardcoded `VITE_API_URL` value.
- Prefer setting `VITE_API_URL` through deployment platform environment variables.

Suggested production checklist:
1. Set secure `JWT_SECRET` and `JWT_REFRESH_SECRET`.
2. Set correct CORS `FRONTEND_URL`.
3. Confirm backend `PORT` and reverse-proxy settings.
4. Ensure MongoDB backup and index strategy.
5. Disable development credentials and seed dependencies in production.

## Additional Documentation

Top-level docs in repository root:
- `DECK_ENGINE_CATEGORY_FEATURE.md`
- `SIGNATURE_FEATURE_SUMMARY.md`
- `HOW_TO_ADD_SIGNATURE.md`
- `SIGNATURE_CODE_EXAMPLES.md`
- `SIGNATURE_UPLOAD_IMPLEMENTATION.md`

## Known Notes and Gotchas

1. `FrontEnd/README.md` previously contained backend-only API docs; this file now documents the full project.
2. Backend `.env.example` uses `PORT=5000`, but runtime fallback in code is `8080`.
3. `FrontEnd/vercel.json` currently hardcodes a production API URL; keep this aligned with your active backend environment.
4. CORS currently logs blocked origins and still allows them; tighten this behavior for strict production environments.
