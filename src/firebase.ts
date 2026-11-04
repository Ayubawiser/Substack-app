import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously, onAuthStateChanged, type User } from "firebase/auth";
import { doc, getDoc, getFirestore, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyA6JphHAmdeS51OJSXC5U-us-ysaPyHsWI",
  authDomain: "subtrack-app1.firebaseapp.com",
  projectId: "subtrack-app1",
  storageBucket: "subtrack-app1.firebasestorage.app",
  messagingSenderId: "412168549482",
  appId: "1:412168549482:web:47a1bca95cb8f4427e4613"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

export async function ensureFirebaseUser(): Promise<User> {
  if (auth.currentUser) return auth.currentUser;
  await signInAnonymously(auth);
  if (auth.currentUser) return auth.currentUser;
  return new Promise((resolve, reject) => {
    const unsubscribe = onAuthStateChanged(auth, user => {
      unsubscribe();
      if (user) resolve(user);
      else reject(new Error("Anonymous sign-in did not return a user."));
    }, reject);
  });
}

const workspaceRef = (uid: string) => doc(db, "users", uid, "appData", "workspace");

export async function loadWorkspace(uid: string): Promise<{ subs?: unknown[]; projects?: unknown[] } | null> {
  const snapshot = await getDoc(workspaceRef(uid));
  return snapshot.exists() ? snapshot.data() as { subs?: unknown[]; projects?: unknown[] } : null;
}

export async function saveWorkspace(uid: string, subs: unknown[], projects: unknown[]): Promise<void> {
  await setDoc(workspaceRef(uid), { subs, projects, updatedAt: new Date().toISOString() }, { merge: true });
}
