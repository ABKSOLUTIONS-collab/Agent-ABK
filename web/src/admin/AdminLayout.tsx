import { NavLink, Outlet, Link } from 'react-router-dom';
import { useAuth } from '../lib/auth-context';

const tabs = [
  { to: '/admin', label: 'Members', end: true },
  { to: '/admin/settings', label: 'Settings' },
  { to: '/admin/logs', label: 'Errors & Health' },
];

export function AdminLayout() {
  const { user, logout } = useAuth();

  return (
    <div className="min-h-screen bg-white">
      <header className="border-b border-neutral-200">
        <div className="max-w-4xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <Link to="/admin" className="text-lg font-semibold text-neutral-900">
              Admin settings
            </Link>
            <p className="text-xs text-neutral-500">{user?.email}</p>
          </div>
          <div className="flex items-center gap-3">
            <Link to="/" className="text-sm text-neutral-500 hover:text-neutral-800">
              Back to app
            </Link>
            <button
              onClick={() => logout()}
              className="rounded-full border border-neutral-200 px-4 py-1.5 text-sm font-medium text-neutral-800 hover:bg-neutral-50"
            >
              Log out
            </button>
          </div>
        </div>
        <nav className="max-w-4xl mx-auto px-6 flex gap-6">
          {tabs.map((tab) => (
            <NavLink
              key={tab.to}
              to={tab.to}
              end={tab.end}
              className={({ isActive }) =>
                `pb-3 text-sm font-medium border-b-2 -mb-px ${
                  isActive ? 'border-neutral-900 text-neutral-900' : 'border-transparent text-neutral-500 hover:text-neutral-800'
                }`
              }
            >
              {tab.label}
            </NavLink>
          ))}
        </nav>
      </header>
      <main className="max-w-4xl mx-auto px-6 py-8">
        <Outlet />
      </main>
    </div>
  );
}
