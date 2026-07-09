import { Link } from 'react-router-dom';
import { useAuth } from '../lib/auth-context';

export function Home() {
  const { user, logout } = useAuth();

  return (
    <div className="min-h-screen bg-white flex flex-col items-center justify-center px-4 text-center">
      <h1 className="text-2xl font-semibold text-neutral-900 mb-2">Welcome, {user?.email}</h1>
      <p className="text-sm text-neutral-500 mb-8">You're signed in.</p>
      <div className="flex gap-3">
        {user?.role === 'admin' && (
          <Link
            to="/admin"
            className="rounded-full border border-neutral-200 px-5 py-2 text-sm font-medium text-neutral-800 hover:bg-neutral-50"
          >
            Admin settings
          </Link>
        )}
        <button
          onClick={() => logout()}
          className="rounded-full bg-neutral-900 px-5 py-2 text-sm font-medium text-white hover:bg-neutral-800"
        >
          Log out
        </button>
      </div>
    </div>
  );
}
