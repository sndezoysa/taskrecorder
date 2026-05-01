// Firebase Configuration
const firebaseConfig = {
    apiKey: "AIzaSyB8VKEBRGybHfgRJJTT2MlBl4kJK8y1PX8",
    authDomain: "taskrecorder-9e9c3.firebaseapp.com",
    projectId: "taskrecorder-9e9c3",
    storageBucket: "taskrecorder-9e9c3.firebasestorage.app",
    messagingSenderId: "416679395956",
    appId: "1:416679395956:web:00755cd5f74785b5aa4721"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);

// Initialize Firestore
const db = firebase.firestore();

console.log("Firebase connected successfully!");
