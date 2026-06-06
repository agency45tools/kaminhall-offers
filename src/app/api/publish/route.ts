import { NextRequest, NextResponse } from 'next/server'

const GITHUB_TOKEN = process.env.GITHUB_TOKEN!
const GITHUB_OWNER = process.env.GITHUB_OWNER ?? 'agency45tools'
const GITHUB_REPO = process.env.GITHUB_REPO ?? 'kaminhall-feed'
const GITHUB_FILE_PATH = 'offers.json'
const GITHUB_BRANCH = process.env.GITHUB_BRANCH ?? 'main'

export async function POST(req: NextRequest) {
  try {
    if (!GITHUB_TOKEN) {
      return NextResponse.json({ error: 'GITHUB_TOKEN not configured' }, { status: 500 })
    }

    const { content } = await req.json()
    if (!content) {
      return NextResponse.json({ error: 'Missing content' }, { status: 400 })
    }

    const apiBase = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE_PATH}`
    const headers = {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
    }

    // Get current file SHA (needed for update)
    let sha: string | undefined
    const getRes = await fetch(`${apiBase}?ref=${GITHUB_BRANCH}`, { headers })
    if (getRes.ok) {
      const data = await getRes.json()
      sha = data.sha
    } else if (getRes.status !== 404) {
      const err = await getRes.text()
      return NextResponse.json({ error: `GitHub GET failed: ${err}` }, { status: 502 })
    }

    // Encode content to base64
    const base64 = Buffer.from(JSON.stringify(content, null, 2)).toString('base64')
    const now = new Date().toISOString()

    const putBody = {
      message: `offers: update ${now}`,
      content: base64,
      branch: GITHUB_BRANCH,
      ...(sha ? { sha } : {}),
    }

    const putRes = await fetch(apiBase, {
      method: 'PUT',
      headers,
      body: JSON.stringify(putBody),
    })

    if (!putRes.ok) {
      const err = await putRes.text()
      return NextResponse.json({ error: `GitHub PUT failed: ${err}` }, { status: 502 })
    }

    const result = await putRes.json()
    return NextResponse.json({
      success: true,
      sha: result.content?.sha,
      commitUrl: result.commit?.html_url,
    })
  } catch (err) {
    console.error('Publish error:', err)
    return NextResponse.json({ error: String(err) }, { status: 500 })
  }
}
