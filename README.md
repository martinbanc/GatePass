# Padel Connect – MVP Starter Guide

This document walks you through creating the **Padel Connect** minimum viable product using the stack you described (Next.js 14, Prisma + SQLite, NextAuth, TailwindCSS, Stripe-ready checkout, and QR code ticketing). The guide is written for beginners using Visual Studio Code and assumes no prior Next.js experience.

## 1. Prerequisites

1. Install [Node.js 20 LTS](https://nodejs.org/en/download) which includes `npm`.
2. Install [Visual Studio Code](https://code.visualstudio.com/Download).
3. (Optional) Install [Git](https://git-scm.com/downloads) if you plan to track changes or deploy via GitHub.

Open VS Code once the tools above are ready.

## 2. Create the project

```bash
npx create-next-app@latest padel-connect \
  --typescript \
  --tailwind \
  --eslint \
  --app \
  --src-dir
```

Move into the new folder and open it in VS Code:

```bash
cd padel-connect
code .
```

> **Tip:** In VS Code, the built-in terminal (``Ctrl + ` ``) lets you run the commands below without leaving the editor.

## 3. Install additional dependencies

Add the packages used by the starter:

```bash
npm install @prisma/client next-auth qrcode stripe
npm install -D prisma tsx
```

## 4. Configure Prisma + SQLite

1. Create a `.env` file in the project root:

   ```bash
   cp .env.local .env
   echo "DATABASE_URL=file:./dev.db" >> .env
   ```

2. Replace the generated `prisma/schema.prisma` with the data model from the concept (Users, Events, Tickets).
3. Run the Prisma commands to generate the client and create the SQLite database:

   ```bash
   npx prisma generate
   npx prisma migrate dev --name init
   ```

4. (Optional) Create a `prisma/seed.ts` file to populate sample events, then run:

   ```bash
   npx tsx prisma/seed.ts
   ```

## 5. Environment variables

Populate the `.env` file with the configuration listed in the concept (NextAuth, email server, and optional Stripe keys). For local development you can leave Stripe keys blank to use the mock checkout path.

## 6. Project structure

Recreate the folders and files from the concept inside the `src` directory. The key pieces are:

- `app` directory with layouts and pages for marketing, events, event detail, and tickets
- `app/api` routes for auth, events, checkout, tickets, and webhooks
- `components` (Header, Footer, EventCard, QR)
- `lib` utilities for Prisma, authentication, payments, and shared helpers
- `public/manifest.json` plus PWA icons

Copy the code snippets from the concept into the corresponding files. VS Code offers multi-file editing and IntelliSense to help paste and adjust the modules quickly.

## 7. Styling with Tailwind CSS

Replace `src/app/globals.css` with the Tailwind utilities shown in the concept. Tailwind classes provide the dark theme, cards, buttons, and layout styles for the MVP.

## 8. Running the development server

Once the files are in place, start the Next.js development server:

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser. You can “Install App” thanks to the included PWA manifest, browse events, and run the mock checkout flow that issues QR-coded tickets.

## 9. Next steps

- Gate checkout behind a NextAuth session to bind tickets to real users.
- Expand the Tickets page to fetch the current user’s tickets via a dedicated API route.
- Switch `DATABASE_URL` to a hosted Postgres instance when deploying to Vercel, then run `npx prisma migrate deploy`.
- Replace the mock checkout with live Stripe by setting the secret keys and implementing signature verification in the webhook.
- Extend the Prisma schema to support clubs, coaches, tournaments, and bookings as outlined in the concept roadmap.

With these steps you have a working Padel Connect MVP and a clear path to iterate on future features.