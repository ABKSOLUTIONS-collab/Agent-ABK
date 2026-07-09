import { useState, type FormEvent } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { api, ApiError } from '../lib/api';
import { AuthCard, Field, PrimaryButton, ErrorText } from '../components/ui';

export function ResetPassword() {
  const [searchParams] = useSearchParams();
  const token = searchParams.get('token') ?? '';
  const navigate = useNavigate();

  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError('');

    if (password.length < 8) {
      setError('Password must be at least 8 characters.');
      return;
    }
    if (password !== confirm) {
      setError('Passwords do not match.');
      return;
    }
    if (!token) {
      setError('This reset link is missing its token.');
      return;
    }

    setSubmitting(true);
    try {
      await api.post('/api/auth/reset-password', { token, newPassword: password });
      navigate('/login', { replace: true, state: { resetComplete: true } });
    } catch (err) {
      setError(err instanceof ApiError && err.status === 400 ? 'This link is invalid or has expired.' : 'Something went wrong.');
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <AuthCard title="Set a new password">
      <form onSubmit={handleSubmit}>
        <ErrorText>{error}</ErrorText>
        <Field
          label="New password"
          type="password"
          autoComplete="new-password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />
        <Field
          label="Confirm new password"
          type="password"
          autoComplete="new-password"
          required
          value={confirm}
          onChange={(e) => setConfirm(e.target.value)}
        />
        <PrimaryButton type="submit" disabled={submitting}>
          {submitting ? 'Saving…' : 'Set new password'}
        </PrimaryButton>
      </form>
      <p className="mt-6 text-sm text-center text-neutral-500">
        <Link to="/login" className="text-neutral-800 font-medium hover:underline">
          Back to login
        </Link>
      </p>
    </AuthCard>
  );
}
