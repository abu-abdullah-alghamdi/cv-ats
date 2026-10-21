/*
 * Firebase connection: passwordless email-link sign in + one Firestore document per user (cvs/{uid}).
 * The page talks to this file through window.CVCloud and the "cvcloud" event.
 * When firebase-config.js still has the placeholder values, the site runs in local mode (browser storage only).
 */
import { firebaseConfig } from './firebase-config.js';

const SDK = '10.14.1';
const EMAIL_KEY = 'cv-ats-email-for-signin';

function ready(api) {
  window.CVCloud = api;
  window.dispatchEvent(new CustomEvent('cvcloud', { detail: api }));
}

const configured = !!(firebaseConfig && firebaseConfig.apiKey && !/^YOUR_/i.test(firebaseConfig.apiKey));

if (!configured) {
  ready({ configured: false });
} else {
  try {
    const [{ initializeApp }, A, F] = await Promise.all([
      import(`https://www.gstatic.com/firebasejs/${SDK}/firebase-app.js`),
      import(`https://www.gstatic.com/firebasejs/${SDK}/firebase-auth.js`),
      import(`https://www.gstatic.com/firebasejs/${SDK}/firebase-firestore.js`)
    ]);
    const app = initializeApp(firebaseConfig);
    const auth = A.getAuth(app);
    auth.languageCode = 'ar';
    const db = F.getFirestore(app);

    // Local testing against the Firebase emulators: run localStorage.setItem('cv-ats-emulators', '1') in the console.
    let useEmu = false;
    try { useEmu = localStorage.getItem('cv-ats-emulators') === '1'; } catch (e) {}
    if (useEmu) {
      A.connectAuthEmulator(auth, 'http://127.0.0.1:9099', { disableWarnings: true });
      F.connectFirestoreEmulator(db, '127.0.0.1', 8080);
    }

    const store = {
      get() { try { return localStorage.getItem(EMAIL_KEY) || ''; } catch (e) { return ''; } },
      set(v) { try { localStorage.setItem(EMAIL_KEY, v); } catch (e) {} },
      del() { try { localStorage.removeItem(EMAIL_KEY); } catch (e) {} }
    };
    const cleanUrl = () => { try { history.replaceState(null, '', location.pathname + location.hash); } catch (e) {} };

    ready({
      configured: true,
      onUser(cb) { return A.onAuthStateChanged(auth, cb); },
      async sendLink(email) {
        await A.sendSignInLinkToEmail(auth, email, {
          url: location.origin + location.pathname,
          handleCodeInApp: true
        });
        store.set(email);
      },
      isLinkInUrl() { return A.isSignInWithEmailLink(auth, location.href); },
      storedEmail() { return store.get(); },
      async completeLink(email) {
        const cred = await A.signInWithEmailLink(auth, email, location.href);
        store.del();
        cleanUrl();
        return cred.user;
      },
      clearLink() { cleanUrl(); },
      signOut() { return A.signOut(auth); },
      async load(uid) {
        const snap = await F.getDoc(F.doc(db, 'cvs', uid));
        return snap.exists() ? snap.data() : null;
      },
      save(uid, data) {
        return F.setDoc(F.doc(db, 'cvs', uid), { ...data, updatedAt: F.serverTimestamp() }, { merge: true });
      }
    });
  } catch (e) {
    console.error('Firebase failed to start', e);
    ready({ configured: false, error: true });
  }
}
