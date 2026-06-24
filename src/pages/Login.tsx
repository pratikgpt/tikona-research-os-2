import { useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { toast } from 'sonner';
import { useAuth } from '@/contexts/AuthContext';
import { Spinner } from '@/components/ui/spinner';

function GoogleIcon() {
  return (
    <svg className="h-4 w-4" viewBox="0 0 24 24">
      <path
        fill="#4285F4"
        d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
      />
      <path
        fill="#34A853"
        d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
      />
      <path
        fill="#FBBC05"
        d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"
      />
      <path
        fill="#EA4335"
        d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"
      />
    </svg>
  );
}

export default function Login() {
  const { user, isLoading, error, signInWithGoogle, clearError } = useAuth();
  const navigate = useNavigate();

  useEffect(() => {
    if (error) {
      toast.error(error);
      clearError();
    }
  }, [error, clearError]);

  useEffect(() => {
    if (user && !isLoading) {
      navigate('/admin', { replace: true });
    }
  }, [user, isLoading, navigate]);

  const handleGoogleSignIn = async () => {
    try {
      await signInWithGoogle();
    } catch {
      toast.error('Failed to initiate sign in');
    }
  };

  return (
    <div className="flex min-h-screen bg-white">
      {/* Left panel - branding */}
      <div className="hidden lg:flex lg:w-1/2 bg-gradient-to-br from-neutral-950 via-neutral-900 to-accent-950 text-white flex-col justify-between p-12 relative overflow-hidden">
        {/* Subtle decorative gradient orb */}
        <div className="absolute -bottom-32 -right-32 h-96 w-96 rounded-full bg-accent-600/10 blur-3xl" />
        <div className="absolute top-1/4 -left-16 h-64 w-64 rounded-full bg-accent-500/5 blur-3xl" />

        <div className="relative">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-white">
              <img src="/tikona-logo.png" alt="Tikona Capital" className="h-8 w-8 object-contain" />
            </div>
            <span className="text-sm font-semibold tracking-tight">Tikona Capital</span>
          </div>
        </div>

        <div className="relative">
          <h1 className="text-5xl font-semibold leading-tight tracking-tight">
            Research OS
          </h1>
          <p className="mt-4 text-xl text-neutral-400 max-w-md leading-relaxed">
            AI-powered equity research platform. Generate comprehensive reports, manage your investment universe, and streamline your analysis workflow.
          </p>
        </div>

        <p className="relative text-xs text-neutral-600">
          &copy; {new Date().getFullYear()} Tikona Capital
        </p>
      </div>

      {/* Right panel - login form */}
      <div className="flex flex-1 items-center justify-center px-6">
        <div className="w-full max-w-sm">
          {/* Mobile logo */}
          <div className="lg:hidden mb-10">
            <div className="flex items-center gap-3">
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-white">
                <img src="/tikona-logo.png" alt="Tikona Capital" className="h-7 w-7 object-contain" />
              </div>
              <span className="text-sm font-semibold text-neutral-900">Tikona Capital</span>
            </div>
          </div>

          <h2 className="text-3xl font-semibold text-neutral-900 tracking-tight">
            Sign in
          </h2>
          <p className="mt-2 text-sm text-neutral-500">
            Continue to your research dashboard
          </p>

          <button
            onClick={handleGoogleSignIn}
            disabled={isLoading}
            className="mt-8 flex h-12 w-full items-center justify-center gap-3 rounded-xl border border-neutral-200 bg-white text-sm font-medium text-neutral-700 shadow-sm transition-all duration-200 hover:shadow-md hover:border-neutral-300 active:scale-[0.98] disabled:opacity-50 disabled:pointer-events-none"
          >
            {isLoading ? (
              <>
                <Spinner size="sm" />
                Signing in...
              </>
            ) : (
              <>
                <GoogleIcon />
                Continue with Google
              </>
            )}
          </button>

          <p className="mt-8 text-xs text-neutral-400">
            By signing in, you agree to our terms of service and privacy policy.
          </p>
        </div>
      </div>
    </div>
  );
}
