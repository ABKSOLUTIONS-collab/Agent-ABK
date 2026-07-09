import { useEffect, useState, type FormEvent } from 'react';
import { api, ApiError } from '../lib/api';

interface Member {
  email: string;
  role: 'admin' | 'user';
  createdAt: number;
  isActive: boolean;
}

export function Members() {
  const [members, setMembers] = useState<Member[] | null>(null);
  const [error, setError] = useState('');
  const [showAdd, setShowAdd] = useState(false);

  function load() {
    api
      .get<Member[]>('/api/admin/users')
      .then(setMembers)
      .catch(() => setError('Failed to load members.'));
  }

  useEffect(load, []);

  async function updateRole(email: string, role: Member['role']) {
    setError('');
    try {
      await api.patch(`/api/admin/users/${encodeURIComponent(email)}`, { role });
      load();
    } catch (err) {
      setError(err instanceof ApiError ? err.message : 'Failed to update role.');
    }
  }

  async function toggleActive(member: Member) {
    setError('');
    try {
      await api.patch(`/api/admin/users/${encodeURIComponent(member.email)}`, { isActive: !member.isActive });
      load();
    } catch (err) {
      setError(err instanceof ApiError ? err.message : 'Failed to update member.');
    }
  }

  async function removeMember(email: string) {
    if (!confirm(`Remove ${email}? This cannot be undone.`)) return;
    setError('');
    try {
      await api.delete(`/api/admin/users/${encodeURIComponent(email)}`);
      load();
    } catch (err) {
      setError(err instanceof ApiError ? err.message : 'Failed to remove member.');
    }
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h2 className="text-base font-semibold text-neutral-900">Members</h2>
        <button
          onClick={() => setShowAdd((v) => !v)}
          className="rounded-full bg-neutral-900 px-4 py-1.5 text-sm font-medium text-white hover:bg-neutral-800"
        >
          {showAdd ? 'Cancel' : 'Add member'}
        </button>
      </div>

      {error && <p className="text-sm text-red-600 mb-4">{error}</p>}

      {showAdd && (
        <AddMemberForm
          onAdded={() => {
            setShowAdd(false);
            load();
          }}
          onError={setError}
        />
      )}

      <table className="w-full text-sm">
        <thead>
          <tr className="text-left text-xs text-neutral-500 border-b border-neutral-200">
            <th className="py-2 font-medium">Email</th>
            <th className="py-2 font-medium">Role</th>
            <th className="py-2 font-medium">Status</th>
            <th className="py-2 font-medium"></th>
          </tr>
        </thead>
        <tbody>
          {members?.map((m) => (
            <tr key={m.email} className="border-b border-neutral-100">
              <td className="py-3 text-neutral-900">{m.email}</td>
              <td className="py-3">
                <select
                  value={m.role}
                  onChange={(e) => updateRole(m.email, e.target.value as Member['role'])}
                  className="rounded-lg border border-neutral-200 px-2 py-1 text-sm"
                >
                  <option value="user">User</option>
                  <option value="admin">Admin</option>
                </select>
              </td>
              <td className="py-3">
                <button
                  onClick={() => toggleActive(m)}
                  className={`rounded-full px-2.5 py-0.5 text-xs font-medium ${
                    m.isActive ? 'bg-green-50 text-green-700' : 'bg-neutral-100 text-neutral-500'
                  }`}
                >
                  {m.isActive ? 'Active' : 'Disabled'}
                </button>
              </td>
              <td className="py-3 text-right">
                <button onClick={() => removeMember(m.email)} className="text-xs text-red-600 hover:underline">
                  Remove
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {members?.length === 0 && <p className="text-sm text-neutral-500 mt-4">No members yet.</p>}
    </div>
  );
}

function AddMemberForm({ onAdded, onError }: { onAdded: () => void; onError: (msg: string) => void }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [role, setRole] = useState<'admin' | 'user'>('user');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    setSubmitting(true);
    onError('');
    try {
      await api.post('/api/admin/users', { email, password, role });
      onAdded();
    } catch (err) {
      onError(err instanceof ApiError ? err.message : 'Failed to add member.');
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <form onSubmit={handleSubmit} className="mb-6 rounded-xl border border-neutral-200 p-4 flex flex-wrap items-end gap-3">
      <label className="text-left">
        <span className="text-xs font-medium text-neutral-500 block mb-1">Email</span>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          className="rounded-lg border border-neutral-200 px-3 py-1.5 text-sm"
        />
      </label>
      <label className="text-left">
        <span className="text-xs font-medium text-neutral-500 block mb-1">Temporary password</span>
        <input
          type="text"
          required
          minLength={8}
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          className="rounded-lg border border-neutral-200 px-3 py-1.5 text-sm"
        />
      </label>
      <label className="text-left">
        <span className="text-xs font-medium text-neutral-500 block mb-1">Role</span>
        <select
          value={role}
          onChange={(e) => setRole(e.target.value as 'admin' | 'user')}
          className="rounded-lg border border-neutral-200 px-3 py-1.5 text-sm"
        >
          <option value="user">User</option>
          <option value="admin">Admin</option>
        </select>
      </label>
      <button
        type="submit"
        disabled={submitting}
        className="rounded-full bg-neutral-900 px-4 py-1.5 text-sm font-medium text-white hover:bg-neutral-800 disabled:opacity-50"
      >
        {submitting ? 'Adding…' : 'Add'}
      </button>
    </form>
  );
}
