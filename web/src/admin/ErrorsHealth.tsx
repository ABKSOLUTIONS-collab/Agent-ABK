import { useEffect, useState } from 'react';
import { api } from '../lib/api';

interface ErrorEntry {
  timestamp: number;
  source: string;
  message: string;
  detail?: string;
}

interface LogsResponse {
  errors: ErrorEntry[];
  health: { status: string; server: string; cachedTools: number; activeUsers: number };
}

export function ErrorsHealth() {
  const [data, setData] = useState<LogsResponse | null>(null);

  useEffect(() => {
    api.get<LogsResponse>('/api/admin/logs').then(setData);
  }, []);

  if (!data) return null;

  return (
    <div>
      <h2 className="text-base font-semibold text-neutral-900 mb-4">Health</h2>
      <div className="grid grid-cols-3 gap-4 mb-8">
        <StatCard label="Status" value={data.health.status} />
        <StatCard label="Cached tools" value={String(data.health.cachedTools)} />
        <StatCard label="Active users" value={String(data.health.activeUsers)} />
      </div>

      <h2 className="text-base font-semibold text-neutral-900 mb-4">Recent errors</h2>
      {data.errors.length === 0 ? (
        <p className="text-sm text-neutral-500">No errors logged.</p>
      ) : (
        <div className="space-y-3">
          {data.errors.map((e, i) => (
            <div key={i} className="rounded-xl border border-neutral-200 p-3">
              <div className="flex items-center justify-between mb-1">
                <span className="text-xs font-medium text-neutral-500">{e.source}</span>
                <span className="text-xs text-neutral-400">{new Date(e.timestamp).toLocaleString()}</span>
              </div>
              <p className="text-sm text-neutral-900">{e.message}</p>
              {e.detail && <pre className="text-xs text-neutral-500 mt-1 whitespace-pre-wrap break-all">{e.detail}</pre>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl border border-neutral-200 p-4">
      <p className="text-xs text-neutral-500 mb-1">{label}</p>
      <p className="text-lg font-semibold text-neutral-900">{value}</p>
    </div>
  );
}
