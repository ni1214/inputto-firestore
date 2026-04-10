import { initializeApp } from 'firebase/app';
import { initializeFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyBDrBUN2elbCAdxfbnTFWQNWF4xhz9yaJ0',
  authDomain: 'kategu-sys-v15.firebaseapp.com',
  projectId: 'kategu-sys-v15',
  storageBucket: 'kategu-sys-v15.firebasestorage.app',
  messagingSenderId: '992448511434',
  appId: '1:992448511434:web:ef53560b55264f1e656333'
};

const app = initializeApp(firebaseConfig);

export const db = initializeFirestore(app, {
  ignoreUndefinedProperties: true
});

export { app, firebaseConfig };
