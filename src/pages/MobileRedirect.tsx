import { useParams } from 'react-router-dom';
import { useEffect } from 'react';
import { Send, Smartphone, Download, ExternalLink } from 'lucide-react';

export default function MobileRedirect() {
  const { id } = useParams<{ id: string }>();

  useEffect(() => {
    if (id) {
      // Automatically attempt to redirect to the mobile app deep link
      const timer = setTimeout(() => {
        window.location.href = `tikonamobile://report/${id}`;
      }, 500);
      return () => clearTimeout(timer);
    }
  }, [id]);

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-950 text-slate-100 p-6 relative overflow-hidden font-sans">
      {/* Background gradients for premium glassmorphism feel */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-accent-600/10 rounded-full blur-3xl" />
      <div className="absolute bottom-1/4 right-1/4 w-96 h-96 bg-purple-600/10 rounded-full blur-3xl" />

      <div className="max-w-md w-full relative z-10 backdrop-blur-md bg-slate-900/60 border border-slate-800 rounded-3xl p-8 shadow-2xl text-center flex flex-col items-center">
        {/* Animated Icon */}
        <div className="relative mb-6">
          <div className="absolute inset-0 bg-accent-500/20 rounded-full blur-xl animate-pulse" />
          <div className="h-16 w-16 rounded-2xl bg-gradient-to-tr from-accent-600 to-purple-600 flex items-center justify-center shadow-lg relative border border-accent-400/30">
            <Smartphone className="h-8 w-8 text-white animate-bounce" />
          </div>
        </div>

        <h2 className="text-2xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-white via-slate-100 to-slate-300 tracking-tight">
          Opening Report
        </h2>
        
        <p className="text-sm text-slate-400 mt-3 leading-relaxed">
          We are redirecting you to the <strong className="text-accent-400 font-semibold">Tikona Research</strong> mobile app to securely view this report.
        </p>

        {/* Loading Spinner */}
        <div className="flex items-center gap-2 mt-6 mb-4 px-4 py-2 rounded-full bg-slate-800/40 border border-slate-800 text-xs font-mono text-slate-400">
          <span className="h-2 w-2 rounded-full bg-accent-500 animate-ping" />
          Securing connection & launching app...
        </div>

        {/* Action Buttons */}
        <div className="w-full mt-6 space-y-3">
          <a
            href={`tikonamobile://report/${id}`}
            className="w-full inline-flex justify-center items-center gap-2 rounded-xl bg-gradient-to-r from-accent-600 to-accent-500 hover:from-accent-500 hover:to-accent-400 px-5 py-3.5 text-sm font-semibold text-white shadow-lg shadow-accent-600/20 transition-all duration-200 active:scale-[0.98]"
          >
            Open in App
            <ExternalLink className="h-4 w-4" />
          </a>

          <div className="pt-6 border-t border-slate-800/80 w-full mt-6">
            <p className="text-xs text-slate-500 mb-4">
              Don't have the mobile app installed yet?
            </p>
            <div className="flex gap-3 justify-center">
              <a
                href="https://play.google.com/store"
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1.5 px-4 py-2.5 rounded-xl border border-slate-800 bg-slate-900 hover:bg-slate-800 text-xs font-semibold text-slate-300 transition-colors"
              >
                <Download className="h-3.5 w-3.5" />
                Google Play
              </a>
              <a
                href="https://apps.apple.com"
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1.5 px-4 py-2.5 rounded-xl border border-slate-800 bg-slate-900 hover:bg-slate-800 text-xs font-semibold text-slate-300 transition-colors"
              >
                <Download className="h-3.5 w-3.5" />
                App Store
              </a>
            </div>
          </div>
        </div>
      </div>

      <div className="absolute bottom-6 text-center w-full left-0 text-slate-600 text-[10px] uppercase tracking-widest font-semibold">
        © Tikona Capital • Secure Gateway
      </div>
    </div>
  );
}
