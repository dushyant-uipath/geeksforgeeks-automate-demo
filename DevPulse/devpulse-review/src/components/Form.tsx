import { useState, useEffect } from 'react';
import type { ChangeEvent } from 'react';
import { Theme } from '@uipath/coded-action-app';
import { codedActionAppService } from '../uipath';
import './Form.css';

interface FormData {
  recipient: string;
  topic: string;
  subject: string;
  bodyHtml: string;
  note: string;
}

const defaultFormData: FormData = {
  recipient: '',
  topic: '',
  subject: '',
  bodyHtml: '',
  note: '',
};

const isDarkTheme = (theme: Theme): boolean =>
  theme === Theme.Dark || theme === Theme.DarkHighContrast;

interface FormProps {
  onInitTheme: (isDark: boolean) => void;
  darkTheme: boolean;
  onToggleTheme: () => void;
}

function Form({ onInitTheme, darkTheme, onToggleTheme }: FormProps) {
  const [formData, setFormData] = useState<FormData>(defaultFormData);
  const [isReadOnly, setIsReadOnly] = useState(false);

  useEffect(() => {
    codedActionAppService.getTask().then((task) => {
      const merged = task.data
        ? { ...defaultFormData, ...(task.data as Partial<FormData>) }
        : defaultFormData;
      setFormData(merged);
      setIsReadOnly(task.isReadOnly);
      onInitTheme(isDarkTheme(task.theme));
    });
  }, [onInitTheme]);

  const handleTextChange = (e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    if (isReadOnly) return;
    const { name, value } = e.target;
    const updated = { ...formData, [name]: value };
    setFormData(updated);
    codedActionAppService.setTaskData(updated);
  };

  const canAct = !isReadOnly;

  const handleSend = async () => {
    await codedActionAppService.completeTask('Send', formData);
  };
  const handleDiscard = async () => {
    await codedActionAppService.completeTask('Discard', formData);
  };

  return (
    <div className="review-app">
      <header className="review-header">
        <div className="review-header__icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect x="3" y="5" width="18" height="14" rx="2" />
            <path d="M3 7l9 6 9-6" />
          </svg>
        </div>
        <div className="review-header__titles">
          <h1 className="review-header__title">DevPulse Review</h1>
          <p className="review-header__subtitle">
            Read today's issue, then send it or discard it.
          </p>
        </div>
        <div className="review-header__actions">
          {isReadOnly && <span className="review-badge">Read only</span>}
          <button
            type="button"
            className="theme-toggle"
            onClick={onToggleTheme}
            aria-label={darkTheme ? 'Switch to light mode' : 'Switch to dark mode'}
            title={darkTheme ? 'Switch to light mode' : 'Switch to dark mode'}
          >
            {darkTheme ? (
              <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="4" />
                <path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41" />
              </svg>
            ) : (
              <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
              </svg>
            )}
          </button>
        </div>
      </header>

      <div className="form-container form-container--enter">
        <section className="form-section">
          <h2 className="form-title">The issue</h2>
          <div className="form-grid">
            <div className="form-group">
              <label htmlFor="recipient">Sending to</label>
              <input id="recipient" readOnly value={formData.recipient} />
            </div>
            <div className="form-group">
              <label htmlFor="topic">Topic</label>
              <input id="topic" readOnly value={formData.topic} />
            </div>
          </div>
          <div className="form-group form-group--spaced">
            <label htmlFor="subject">Subject</label>
            <input id="subject" readOnly value={formData.subject} />
          </div>
          <div className="form-group form-group--spaced">
            <label htmlFor="preview">Preview</label>
            <iframe
              id="preview"
              className="issue-preview"
              title="Newsletter preview"
              sandbox=""
              srcDoc={formData.bodyHtml}
            />
          </div>
        </section>

        <section className="form-section">
          <h2 className="form-title">Your decision</h2>
          <div className="form-group form-group--spaced">
            <label htmlFor="note">Note (optional)</label>
            <textarea
              id="note"
              name="note"
              rows={3}
              placeholder="Anything the team should know about this issue…"
              value={formData.note}
              onChange={handleTextChange}
              readOnly={isReadOnly}
            />
          </div>
        </section>
      </div>

      <div className="form-buttons">
        <button
          type="button"
          className="outcome-btn outcome-btn--secondary"
          onClick={handleDiscard}
          disabled={!canAct}
        >
          Discard
        </button>
        <button
          type="button"
          className="outcome-btn outcome-btn--primary"
          onClick={handleSend}
          disabled={!canAct}
        >
          Send
        </button>
      </div>
    </div>
  );
}

export default Form;
