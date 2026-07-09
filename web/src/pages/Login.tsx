import { useState, type FormEvent } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import { useAuth } from '../lib/auth-context';
import { AuthCard, Field, PrimaryButton, ErrorText } from '../components/ui';

export function Login() {
  const { login } = useAuth();
  const navigate = useNavigate();
  const location = useLocation();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setError('');
    setSubmitting(true);
    try {
      await login(email, password);
      const from = (location.state as { from?: string } | null)?.from ?? '/';
      navigate(from, { replace: true });
    } catch {
      setError('Incorrect email or password.');
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <AuthCard title="Welcome back">
      <form onSubmit={handleSubmit}>
        <ErrorText>{error}</ErrorText>
        <Field
          label="Email address"
          type="email"
          autoComplete="username"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
        />
        <Field
          label="Password"
          type="password"
          autoComplete="current-password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />
        <div className="mb-4 text-right">
          <Link to="/forgot-password" className="text-xs text-neutral-500 hover:text-neutral-800">
            Forgot password?
          </Link>
        </div>
        <PrimaryButton type="submit" disabled={submitting}>
          {submitting ? 'Signing in…' : 'Continue'}
        </PrimaryButton>
      </form>
    </AuthCard>
  );
}
