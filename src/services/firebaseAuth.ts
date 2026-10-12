import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  signInWithRedirect,
  getRedirectResult,
  GoogleAuthProvider,
  onAuthStateChanged,
  User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

export const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

const provider = new GoogleAuthProvider();
SCOPES.forEach((scope) => provider.addScope(scope));
provider.setCustomParameters({
  prompt: 'select_account',
});

// Flag to indicate if we are in the middle of a sign-in flow.
let isSigningIn = false;
// Cache the access token strictly in memory (never in localStorage/sessionStorage).
let cachedAccessToken: string | null = null;

// Initialize auth state listener. Call this on app load.
export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  getRedirectResult(auth)
    .then((result) => {
      if (result) {
        const credential = GoogleAuthProvider.credentialFromResult(result);
        if (credential?.accessToken) {
          cachedAccessToken = credential.accessToken;
          if (onAuthSuccess) onAuthSuccess(result.user, cachedAccessToken);
        }
      }
    })
    .catch((err) => {
      console.warn('Redirect sign-in check:', err);
    });

  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      if (cachedAccessToken) {
        if (onAuthSuccess) onAuthSuccess(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        if (onAuthFailure) onAuthFailure();
      }
    } else {
      cachedAccessToken = null;
      if (onAuthFailure) onAuthFailure();
    }
  });
};

function loadGoogleIdentityServicesScript(): Promise<void> {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined') {
      reject(new Error('Window tidak tersedia'));
      return;
    }
    if ((window as any).google?.accounts?.oauth2) {
      resolve();
      return;
    }
    const existing = document.getElementById('google-gsi-client-script');
    if (existing) {
      existing.addEventListener('load', () => resolve());
      existing.addEventListener('error', () =>
        reject(new Error('Gagal memuat script Google Identity Services'))
      );
      return;
    }
    const script = document.createElement('script');
    script.id = 'google-gsi-client-script';
    script.src = 'https://accounts.google.com/gsi/client';
    script.async = true;
    script.defer = true;
    script.onload = () => resolve();
    script.onerror = () =>
      reject(new Error('Gagal memuat script Google Identity Services'));
    document.head.appendChild(script);
  });
}

async function signInWithGsiFallback(
  clientIdOverride?: string
): Promise<{ user: User; accessToken: string }> {
  await loadGoogleIdentityServicesScript();
  const clientId = (clientIdOverride || firebaseConfig.oAuthClientId || '').trim();
  if (!clientId) {
    throw new Error('OAuth Client ID tidak ditemukan.');
  }

  return new Promise((resolve, reject) => {
    try {
      const tokenClient = (window as any).google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: `${SCOPES.join(' ')} openid email profile`,
        prompt: 'select_account',
        callback: async (tokenResponse: any) => {
          if (tokenResponse?.error) {
            reject(
              new Error(
                tokenResponse.error_description ||
                  tokenResponse.error ||
                  'Gagal otorisasi Google OAuth.'
              )
            );
            return;
          }
          const accessToken = tokenResponse?.access_token;
          if (!accessToken) {
            reject(new Error('Token akses tidak diterima dari Google.'));
            return;
          }
          cachedAccessToken = accessToken;

          // Fetch basic user profile for display
          let email = 'Admin Terhubung (Google OAuth)';
          let displayName = 'Administrator';
          try {
            const infoRes = await fetch(
              'https://www.googleapis.com/oauth2/v3/userinfo',
              {
                headers: { Authorization: `Bearer ${accessToken}` },
              }
            );
            if (infoRes.ok) {
              const info = await infoRes.json();
              email = info.email || email;
              displayName = info.name || displayName;
            }
          } catch {
            // Ignore profile fetch error
          }

          const pseudoUser = {
            email,
            displayName,
            uid: 'gsi-oauth-user',
          } as unknown as User;

          resolve({ user: pseudoUser, accessToken });
        },
        error_callback: (err: any) => {
          reject(
            new Error(
              err?.message ||
                'Popup Google ditutup atau domain belum diizinkan di Google Cloud Console.'
            )
          );
        },
      });
      tokenClient.requestAccessToken();
    } catch (e) {
      reject(e);
    }
  });
}

// Must be called from a button click or user interaction
export const googleSignIn = async (
  customClientId?: string
): Promise<{
  user: User;
  accessToken: string;
} | null> => {
  try {
    isSigningIn = true;

    // If user explicitly provided a custom OAuth Client ID for Vercel, use GIS directly
    if (customClientId && customClientId.trim()) {
      return await signInWithGsiFallback(customClientId.trim());
    }

    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Gagal mendapatkan token akses Google Sheets.');
    }

    cachedAccessToken = credential.accessToken;
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Sign in error:', error);
    const code = error?.code || '';
    const msg = error?.message || '';

    // 1. If domain is unauthorized in Firebase Auth (e.g. deployed on *.vercel.app),
    // try direct Google Identity Services (GIS) popup before reporting domain instructions
    if (
      code === 'auth/unauthorized-domain' ||
      msg.includes('unauthorized-domain')
    ) {
      try {
        return await signInWithGsiFallback(customClientId);
      } catch {
        const host =
          typeof window !== 'undefined' ? window.location.hostname : 'vercel.app';
        const enhancedErr: any = new Error(
          `Domain "${host}" belum didaftarkan pada Firebase Authorized Domains. Tambahkan "${host}" di Firebase Console atau gunakan fitur Koneksi Langsung / Apps Script Bridge di tab Database.`
        );
        enhancedErr.code = 'auth/unauthorized-domain';
        enhancedErr.hostname = host;
        throw enhancedErr;
      }
    }

    // 2. If popup was blocked on mobile/browser, fall back to GIS or redirect flow
    if (
      code === 'auth/popup-blocked' ||
      code === 'auth/cancelled-popup-request'
    ) {
      await signInWithRedirect(auth, provider);
      return null;
    }

    throw error;
  } finally {
    isSigningIn = false;
  }
};

export const getAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken;
};

export const logout = async () => {
  await auth.signOut();
  cachedAccessToken = null;
};

export const FIREBASE_PROJECT_ID = firebaseConfig.projectId;
export const GOOGLE_OAUTH_CLIENT_ID = firebaseConfig.oAuthClientId;


