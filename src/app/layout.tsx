import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'КАМІНХОЛ — Управління пропозиціями',
  description: 'Внутрішній інструмент управління offers.json',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="uk">
      <body>{children}</body>
    </html>
  )
}
