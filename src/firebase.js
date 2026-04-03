/**
 * Firebase Realtime Database integration for Pricing Simulation
 * 
 * Data structure in Firebase:
 *   /games/{gameId}/
 *     meta: { status, currentQuarter, quarterStarted, sdPenaltyEnabled, aipHistory, createdAt }
 *     teams: { teamId: { name, members, pendingPrice, pendingPromo, submitted, quarters, ... } }
 *     config: { allowIndividualPlay, sheetsUrl }
 */

import { initializeApp } from "firebase/app";
import { getDatabase, ref, set, get, update, onValue, off, push } from "firebase/database";

// Firebase config — replace with your project's config
// To set this up:
// 1. Go to https://console.firebase.google.com
// 2. Create a project (or use existing)
// 3. Click "Web" to add a web app
// 4. Copy the firebaseConfig object here
// 5. Enable Realtime Database (test mode for now)
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || "AIzaSyCAhsy4XZECKJoVtPXePHtaCNoj-3J_R2o",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "pricing-simulation-4ceee.firebaseapp.com",
  databaseURL: import.meta.env.VITE_FIREBASE_DATABASE_URL || "https://pricing-simulation-4ceee-default-rtdb.firebaseio.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "pricing-simulation-4ceee",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "pricing-simulation-4ceee.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "386476392838",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:386476392838:web:54a797ad9a2a80bbd12c07",
};

let app = null;
let db = null;

export function isFirebaseConfigured() {
  return !!(firebaseConfig.databaseURL && firebaseConfig.apiKey);
}

function getDb() {
  if (!db) {
    if (!isFirebaseConfigured()) return null;
    app = initializeApp(firebaseConfig);
    db = getDatabase(app);
  }
  return db;
}

// ─── FACULTY: Create/Update Game ──────────────────────────────────────────

export async function createGame(gameId, gameData) {
  const database = getDb();
  if (!database) return false;
  try {
    await set(ref(database, `games/${gameId}`), {
      meta: {
        status: gameData.status || "waiting",
        currentQuarter: gameData.currentQuarter || 1,
        quarterStarted: gameData.quarterStarted || false,
        sdPenaltyEnabled: gameData.sdPenaltyEnabled !== false,
        aipHistory: gameData.aipHistory || [],
        aipInput: gameData.aipInput || 10,
        createdAt: new Date().toISOString(),
      },
      teams: {},
      config: {
        allowIndividualPlay: gameData.allowIndividualPlay !== false,
        sheetsUrl: gameData.sheetsUrl || "",
      },
    });
    return true;
  } catch (e) {
    console.error("createGame error:", e);
    return false;
  }
}

export async function updateGameMeta(gameId, metaUpdates) {
  const database = getDb();
  if (!database) return false;
  try {
    const updates = {};
    Object.keys(metaUpdates).forEach(k => {
      updates[`games/${gameId}/meta/${k}`] = metaUpdates[k];
    });
    await update(ref(database), updates);
    return true;
  } catch (e) {
    console.error("updateGameMeta error:", e);
    return false;
  }
}

export async function saveTeamsToGame(gameId, teams) {
  const database = getDb();
  if (!database) return false;
  try {
    const teamsData = {};
    teams.forEach(t => {
      teamsData[t.id] = {
        name: t.name,
        section: t.section || "",
        members: t.members || [],
        isIndividual: t.isIndividual || false,
        pendingPrice: t.pendingPrice || null,
        pendingPromo: t.pendingPromo || null,
        submitted: t.submitted || false,
        quarters: t.quarters || [],
      };
    });
    await set(ref(database, `games/${gameId}/teams`), teamsData);
    return true;
  } catch (e) {
    console.error("saveTeamsToGame error:", e);
    return false;
  }
}

// ─── STUDENT: Submit price ────────────────────────────────────────────────

export async function submitStudentPrice(gameId, teamId, price, promo, observations, nextStrategy) {
  const database = getDb();
  if (!database) return false;
  try {
    await update(ref(database, `games/${gameId}/teams/${teamId}`), {
      pendingPrice: price,
      pendingPromo: promo || 0,
      submitted: true,
      lastObservations: observations || "",
      lastNextStrategy: nextStrategy || "",
    });
    return true;
  } catch (e) {
    console.error("submitStudentPrice error:", e);
    return false;
  }
}

// ─── STUDENT: Register/join a game ────────────────────────────────────────

export async function joinGame(gameId, teamId, teamName, members, isIndividual) {
  const database = getDb();
  if (!database) return { ok: false, message: "Firebase not configured" };
  try {
    // Check if game exists
    const gameSnap = await get(ref(database, `games/${gameId}/meta`));
    if (!gameSnap.exists()) return { ok: false, message: "Game not found. Check your code." };

    // Register/update team
    await update(ref(database, `games/${gameId}/teams/${teamId}`), {
      name: teamName,
      members: members,
      isIndividual: isIndividual || false,
      pendingPrice: null,
      pendingPromo: null,
      submitted: false,
    });

    return { ok: true, meta: gameSnap.val() };
  } catch (e) {
    console.error("joinGame error:", e);
    return { ok: false, message: e.message };
  }
}

// ─── Check if game exists ─────────────────────────────────────────────────

export async function checkGameExists(gameId) {
  const database = getDb();
  if (!database) return null;
  try {
    const snap = await get(ref(database, `games/${gameId}/meta`));
    return snap.exists() ? snap.val() : null;
  } catch (e) {
    console.error("checkGameExists error:", e);
    return null;
  }
}

// ─── REAL-TIME LISTENERS ──────────────────────────────────────────────────

// Listen to game meta (quarter, status, AIP)
export function listenToGameMeta(gameId, callback) {
  const database = getDb();
  if (!database) return () => {};
  const metaRef = ref(database, `games/${gameId}/meta`);
  const unsubscribe = onValue(metaRef, (snap) => {
    if (snap.exists()) callback(snap.val());
  });
  return () => off(metaRef);
}

// Listen to all teams (faculty view)
export function listenToTeams(gameId, callback) {
  const database = getDb();
  if (!database) return () => {};
  const teamsRef = ref(database, `games/${gameId}/teams`);
  const unsubscribe = onValue(teamsRef, (snap) => {
    if (snap.exists()) {
      const data = snap.val();
      const teams = Object.entries(data).map(([id, t]) => ({ id, ...t }));
      callback(teams);
    }
  });
  return () => off(teamsRef);
}

// Listen to a specific team (student view)
export function listenToTeam(gameId, teamId, callback) {
  const database = getDb();
  if (!database) return () => {};
  const teamRef = ref(database, `games/${gameId}/teams/${teamId}`);
  const unsubscribe = onValue(teamRef, (snap) => {
    if (snap.exists()) callback(snap.val());
  });
  return () => off(teamRef);
}

// ─── FACULTY: Process quarter results ─────────────────────────────────────

export async function processQuarterResults(gameId, updatedTeams, newQuarter, aipHistory, isFinished) {
  const database = getDb();
  if (!database) return false;
  try {
    const updates = {};
    // Update meta
    updates[`games/${gameId}/meta/currentQuarter`] = newQuarter;
    updates[`games/${gameId}/meta/status`] = isFinished ? "finished" : "waiting";
    updates[`games/${gameId}/meta/quarterStarted`] = false;
    updates[`games/${gameId}/meta/aipHistory`] = aipHistory;

    // Update each team's quarters and reset submission state
    updatedTeams.forEach(t => {
      updates[`games/${gameId}/teams/${t.id}/quarters`] = t.quarters;
      updates[`games/${gameId}/teams/${t.id}/pendingPrice`] = null;
      updates[`games/${gameId}/teams/${t.id}/pendingPromo`] = null;
      updates[`games/${gameId}/teams/${t.id}/submitted`] = false;
    });

    await update(ref(database), updates);
    return true;
  } catch (e) {
    console.error("processQuarterResults error:", e);
    return false;
  }
}

export async function startQuarter(gameId, aipInput) {
  return updateGameMeta(gameId, {
    quarterStarted: true,
    status: "active",
    aipInput: aipInput,
  });
}

// Get available team IDs for a game
export async function getGameTeamIds(gameId) {
  const database = getDb();
  if (!database) return [];
  try {
    const snap = await get(ref(database, `games/${gameId}/teams`));
    if (!snap.exists()) return [];
    return Object.entries(snap.val()).map(([id, t]) => ({
      id,
      name: t.name,
      isIndividual: t.isIndividual || false,
      memberCount: (t.members || []).length,
    }));
  } catch (e) {
    return [];
  }
}
