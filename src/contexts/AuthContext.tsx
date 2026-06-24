import {
  createContext,
  useContext,
  useEffect,
  useState,
  useCallback,
  type ReactNode,
} from 'react';
import type { Session, User } from '@supabase/supabase-js';
import { supabase } from '@/lib/supabase';
import { AUDIT_ACTIONS } from '@/lib/constants';

export type UserRole = 'admin';

interface AuthContextType {
  user: User | null;
  session: Session | null;
  role: UserRole | null;
  isLoading: boolean;
  error: string | null;
  signInWithGoogle: () => Promise<void>;
  signOut: () => Promise<void>;
  clearError: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

interface AuthProviderProps {
  children: ReactNode;
}

export function AuthProvider({ children }: AuthProviderProps) {
  const [user, setUser] = useState<User | null>(null);
  const [session, setSession] = useState<Session | null>(null);
  const [role, setRole] = useState<UserRole | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Log audit event
  const logAuditEvent = useCallback(async (action: string, userEmail: string, details: Record<string, unknown> = {}) => {
    try {
      await supabase.from('audit_logs').insert({
        user_email: userEmail,
        action,
        details,
      });
    } catch (err) {
      console.error('Failed to log audit event:', err);
    }
  }, []);

  // All users are admin (admin-only app)
  const determineRole = useCallback((): UserRole => {
    return 'admin';
  }, []);

  // Handle session changes
  const handleSession = useCallback((newSession: Session | null) => {
    if (newSession?.user) {
      setUser(newSession.user);
      setSession(newSession);
      setRole(determineRole());
      setError(null);
    } else {
      setUser(null);
      setSession(null);
      setRole(null);
    }
    setIsLoading(false);
  }, [determineRole]);

  // Initialize auth state
  useEffect(() => {
    let initialHandled = false;

    const { data: { subscription } } = supabase.auth.onAuthStateChange(
      (event, newSession) => {
        console.log('Auth event:', event);

        if (event === 'SIGNED_IN' && newSession?.user.email) {
          logAuditEvent(AUDIT_ACTIONS.USER_LOGIN, newSession.user.email, {
            provider: 'google',
            timestamp: new Date().toISOString(),
          });
        }

        initialHandled = true;
        handleSession(newSession);
      }
    );

    supabase.auth.getSession().then(({ data: { session: initialSession } }) => {
      if (!initialHandled) {
        handleSession(initialSession);
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, [handleSession, logAuditEvent]);

  // Sign in with Google
  const signInWithGoogle = async () => {
    setIsLoading(true);
    setError(null);

    try {
      const { error: signInError } = await supabase.auth.signInWithOAuth({
        provider: 'google',
        options: {
          redirectTo: `${window.location.origin}/login`,
          queryParams: {
            access_type: 'offline',
            prompt: 'consent',
          },
        },
      });

      if (signInError) {
        throw signInError;
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to sign in');
      setIsLoading(false);
    }
  };

  // Sign out
  const signOut = async () => {
    setIsLoading(true);

    try {
      if (user?.email) {
        await logAuditEvent(AUDIT_ACTIONS.USER_LOGOUT, user.email, {
          timestamp: new Date().toISOString(),
        });
      }

      await supabase.auth.signOut();
      setUser(null);
      setSession(null);
      setRole(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to sign out');
    } finally {
      setIsLoading(false);
    }
  };

  const clearError = () => setError(null);

  return (
    <AuthContext.Provider
      value={{
        user,
        session,
        role,
        isLoading,
        error,
        signInWithGoogle,
        signOut,
        clearError,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
}

// eslint-disable-next-line react-refresh/only-export-components
export function useAuth() {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
}
