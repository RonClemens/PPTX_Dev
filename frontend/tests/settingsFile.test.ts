// Run with: npm test   (node's built-in test runner; no extra dependencies)
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { parseSettingsFile } from '../src/settingsFile.ts'

test('parses a .env file with comments, quotes, export and inline comments', () => {
  const r = parseSettingsFile(`
# my work gateway
export ANTHROPIC_API_KEY="sk-work-123"
ANTHROPIC_BASE_URL=https://gw.example.mil/anthropic/   # note: trailing slash
PPTX_DEV_AUTH_MODE=Bearer
PPTX_DEV_EXTRA_HEADERS='X-Tenant-Id: acme\\nX-Other: 1'
PPTX_DEV_AI_MODEL=anthropic.claude-sonnet-4-5-20250929-v1:0
UNRELATED=ignored
`)
  assert.deepEqual(r.settings, {
    apiKey: 'sk-work-123',
    baseUrl: 'https://gw.example.mil/anthropic',
    authMode: 'bearer',
    extraHeaders: 'X-Tenant-Id: acme\nX-Other: 1',
    model: 'anthropic.claude-sonnet-4-5-20250929-v1:0',
  })
  assert.equal(r.found.length, 5)
  assert.deepEqual(r.warnings, [])
})

test('ANTHROPIC_AUTH_TOKEN alone implies Bearer mode (as in Claude Code)', () => {
  const r = parseSettingsFile('ANTHROPIC_AUTH_TOKEN=tok-abc\nANTHROPIC_BASE_URL=https://gw.example')
  assert.equal(r.settings.apiKey, 'tok-abc')
  assert.equal(r.settings.authMode, 'bearer')
})

test('an explicit API key beats the token alias and keeps the default mode', () => {
  const r = parseSettingsFile('ANTHROPIC_API_KEY=sk-1\nANTHROPIC_AUTH_TOKEN=tok-2')
  assert.equal(r.settings.apiKey, 'sk-1')
  assert.equal(r.settings.authMode, undefined)
})

test('parses JSON with camelCase names and header objects', () => {
  const r = parseSettingsFile(
    JSON.stringify({ apiKey: 'k', baseUrl: 'https://gw', authMode: 'both', extraHeaders: { 'X-A': '1', 'X-B': '2' }, model: 'm' }),
  )
  assert.deepEqual(r.settings, { apiKey: 'k', baseUrl: 'https://gw', authMode: 'both', extraHeaders: 'X-A: 1\nX-B: 2', model: 'm' })
})

test('parses Claude Code style {"env": {...}} settings', () => {
  const r = parseSettingsFile(JSON.stringify({ env: { ANTHROPIC_AUTH_TOKEN: 't', ANTHROPIC_BASE_URL: 'https://gw', ANTHROPIC_MODEL: 'mm' }, other: 1 }))
  assert.deepEqual(r.settings, { apiKey: 't', baseUrl: 'https://gw', authMode: 'bearer', model: 'mm' })
})

test('names are case- and separator-insensitive', () => {
  const r = parseSettingsFile('anthropic-api-key=k1\nBase_URL=https://x')
  assert.equal(r.settings.apiKey, 'k1')
  assert.equal(r.settings.baseUrl, 'https://x')
})

test('an unknown auth mode is ignored with a warning, not an error', () => {
  const r = parseSettingsFile('ANTHROPIC_API_KEY=k\nPPTX_DEV_AUTH_MODE=oauth')
  assert.equal(r.settings.authMode, undefined)
  assert.equal(r.warnings.length, 1)
  assert.match(r.warnings[0], /oauth/)
})

test('rejects empty, unrecognized and broken files with readable messages', () => {
  assert.throws(() => parseSettingsFile('   \n'), /empty/)
  assert.throws(() => parseSettingsFile('hello world, not settings'), /No AI settings found/)
  assert.throws(() => parseSettingsFile('{"unrelated": 1}'), /No AI settings found/)
  assert.throws(() => parseSettingsFile('{ broken json'), /JSON/)
})

test('handles Windows line endings and a byte-order mark', () => {
  const r = parseSettingsFile('﻿ANTHROPIC_API_KEY=k\r\nANTHROPIC_BASE_URL=https://gw\r\n')
  assert.equal(r.settings.apiKey, 'k')
  assert.equal(r.settings.baseUrl, 'https://gw')
})

test('never echoes the secret into warnings or found labels', () => {
  const r = parseSettingsFile('ANTHROPIC_API_KEY=super-secret-xyz\nPPTX_DEV_AUTH_MODE=nope')
  assert.ok(!JSON.stringify([r.found, r.warnings]).includes('super-secret-xyz'))
})
