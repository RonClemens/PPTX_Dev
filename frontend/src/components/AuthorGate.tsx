import { useState } from 'react'

interface Props {
  onSubmit: (name: string) => void
}

export default function AuthorGate({ onSubmit }: Props) {
  const [name, setName] = useState('')

  function submit() {
    const trimmed = name.trim()
    if (trimmed) onSubmit(trimmed)
  }

  return (
    <div className="modal-backdrop">
      <div className="modal">
        <div className="modal-header">
          <h3>Welcome</h3>
        </div>
        <p className="modal-hint">
          What's your name? It'll be recorded as the author on every comment and text edit
          you make in this session — visible in PowerPoint just like any other reviewer's comments.
        </p>
        <div className="modal-row">
          <input
            className="modal-input"
            placeholder="Your name"
            value={name}
            autoFocus
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') submit()
            }}
          />
          <button disabled={!name.trim()} onClick={submit}>
            Continue
          </button>
        </div>
      </div>
    </div>
  )
}
