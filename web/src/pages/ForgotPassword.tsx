import { useState, type FormEvent } from 'react';
import { Link } from 'react-router-dom';
import { api } from '../lib/api';
import { AuthCard, Field, PrimaryButton, SuccessText } from '../components/ui';

export function ForgotPassword() {
  const [email, setEmail] = useState('');
  const [sent, setSent] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitting(true);
    try {
      await api.post('/api/auth/forgot-password', { email });
    } finally {
      setSubmitting(false);
      setSent(true);
    }
  }

  return (
    <AuthCard title="Reset your password">
      {sent ? (
        <SuccessText>
          If an account exists for {email}, we've sent a link to reset the password. Check your inbox.
        </SuccessText>
      ) : (
        <form onSubmit={handleSubmit}>
          <p className="text-sm text-neutral-500 mb-4 text-left">
            Enter your email address and we'll send you a link to set a new password.
          </p>
          <Field
            label="Email address"
            type="email"
            autoComplete="username"
            required
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />
          <PrimaryButton type="submit" disabled={submitting}>
            {submitting ? 'Sending…' : 'Send reset link'}
          </PrimaryButton>
        </form>
      )}
      <p className="mt-6 text-sm text-center text-neutral-500">
        <Link to="/login" className="text-neutral-800 font-medium hover:underline">
          Back to login
        </Link>
      </p>
    </AuthCard>
  );
}
