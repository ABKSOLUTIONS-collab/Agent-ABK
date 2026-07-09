import type { ButtonHTMLAttributes, InputHTMLAttributes, ReactNode } from 'react';

export function AuthCard({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="min-h-screen flex items-center justify-center bg-white px-4">
      <div className="w-full max-w-sm">
        <h1 className="text-2xl font-semibold text-center text-neutral-900 mb-6">{title}</h1>
        {children}
      </div>
    </div>
  );
}

export function Field({ label, ...props }: { label: string } & InputHTMLAttributes<HTMLInputElement>) {
  return (
    <label className="block mb-4 text-left">
      <span className="text-xs font-medium text-neutral-500">{label}</span>
      <input
        {...props}
        className="mt-1 w-full rounded-xl border border-neutral-200 px-3.5 py-2.5 text-sm text-neutral-900 outline-none focus:border-neutral-400 focus:ring-1 focus:ring-neutral-300"
      />
    </label>
  );
}

export function PrimaryButton({ children, className = '', ...props }: ButtonHTMLAttributes<HTMLButtonElement>) {
  return (
    <button
      {...props}
      className={`w-full rounded-full bg-neutral-900 text-white text-sm font-medium py-2.5 hover:bg-neutral-800 disabled:opacity-50 disabled:cursor-not-allowed transition-colors ${className}`}
    >
      {children}
    </button>
  );
}

export function ErrorText({ children }: { children: ReactNode }) {
  if (!children) return null;
  return <p className="text-sm text-red-600 mb-4 text-left">{children}</p>;
}

export function SuccessText({ children }: { children: ReactNode }) {
  if (!children) return null;
  return <p className="text-sm text-green-700 mb-4 text-left">{children}</p>;
}
