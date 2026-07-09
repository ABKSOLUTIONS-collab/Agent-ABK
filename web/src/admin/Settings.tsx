import { useEffect, useState, type FormEvent } from 'react';
import { api, ApiError } from '../lib/api';

export function Settings() {
  const [resetEmailSender, setResetEmailSender] = useState('');
  const [loaded, setLoaded] = useState(false);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    api
      .get<{ resetEmailSender: string }>('/api/admin/settings')
      .then((s) => setResetEmailSender(s.resetEmailSender))
      .finally(() => setLoaded(true));
  }, []);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSaving(true);
    setMessage('');
    setError('');
    try {
      await api.put('/api/admin/settings', { resetEmailSender });
      setMessage('Saved.');
    } catch (err) {
      setError(err instanceof ApiError ? err.message : 'Failed to save.');
    } finally {
      setSaving(false);
    }
  }

  if (!loaded) return null;

  return (
    <div className="max-w-md">
      <h2 className="text-base font-semibold text-neutral-900 mb-1">Password reset sender</h2>
      <p className="text-sm text-neutral-500 mb-6">
        The mailbox that "forgot password" emails are sent from. Requires this mailbox to be granted send permission for
        the app's Azure AD registration.
      </p>
      <form onSubmit={handleSubmit}>
        <label className="block mb-4 text-left">
          <span className="text-xs font-medium text-neutral-500">Sender email address</span>
          <input
            type="email"
            required
            value={resetEmailSender}
            onChange={(e) => setResetEmailSender(e.target.value)}
            className="mt-1 w-full rounded-xl border border-neutral-200 px-3.5 py-2.5 text-sm outline-none focus:border-neutral-400 focus:ring-1 focus:ring-neutral-300"
          />
        </label>
        {error && <p className="text-sm text-red-600 mb-4">{error}</p>}
        {message && <p className="text-sm text-green-700 mb-4">{message}</p>}
        <button
          type="submit"
          disabled={saving}
          className="rounded-full bg-neutral-900 px-4 py-1.5 text-sm font-medium text-white hover:bg-neutral-800 disabled:opacity-50"
        >
          {saving ? 'Saving…' : 'Save'}
        </button>
      </form>
    </div>
  );
}
