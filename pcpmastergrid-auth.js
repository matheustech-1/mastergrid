const PCP_AUTH_KEY = "pcpmastergrid_auth_v1";
const PCP_AUTH_HINT_KEY = "pcpmastergrid_auth_hint_v1";
const PCP_AUTH_USERS_KEY = "pcpmastergrid_users_v1";
const PCP_AUTH_MAX_AGE_MS = 1000 * 60 * 60 * 12;

function nowIso() {
  return new Date().toISOString();
}

function safeParse(raw) {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function readSession() {
  return safeParse(localStorage.getItem(PCP_AUTH_KEY));
}

function readUsers() {
  const users = safeParse(localStorage.getItem(PCP_AUTH_USERS_KEY));
  return Array.isArray(users) ? users : [];
}

function writeUsers(users) {
  localStorage.setItem(PCP_AUTH_USERS_KEY, JSON.stringify(users));
}

function rememberHint(userName) {
  if (!userName) return;
  localStorage.setItem(PCP_AUTH_HINT_KEY, userName);
}

function getRememberedHint() {
  return localStorage.getItem(PCP_AUTH_HINT_KEY) || "";
}

function normalizeUserName(userName) {
  return String(userName || "").trim();
}

function findUser(userName) {
  const normalized = normalizeUserName(userName).toLocaleLowerCase("pt-BR");
  return readUsers().find((user) => String(user.userName || "").toLocaleLowerCase("pt-BR") === normalized) || null;
}

function isSessionValid(session) {
  if (!session || !session.userName || !session.loggedInAt) return false;
  const lastAccess = Date.parse(session.lastAccessAt || session.loggedInAt);
  if (Number.isNaN(lastAccess)) return false;
  return (Date.now() - lastAccess) <= PCP_AUTH_MAX_AGE_MS;
}

function writeSession(userName) {
  const session = {
    userName: normalizeUserName(userName),
    loggedInAt: nowIso(),
    lastAccessAt: nowIso(),
    lastPage: "login.html"
  };
  localStorage.setItem(PCP_AUTH_KEY, JSON.stringify(session));
  rememberHint(session.userName);
  return session;
}

function clearSession() {
  localStorage.removeItem(PCP_AUTH_KEY);
}

function touchSession(pageName) {
  const session = readSession();
  if (!isSessionValid(session)) {
    clearSession();
    return null;
  }

  const nextSession = {
    ...session,
    lastAccessAt: nowIso(),
    lastPage: pageName || session.lastPage || ""
  };
  localStorage.setItem(PCP_AUTH_KEY, JSON.stringify(nextSession));
  return nextSession;
}

function formatDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("pt-BR");
}

function createAccount(userName, password) {
  const normalizedUser = normalizeUserName(userName);
  const normalizedPassword = String(password || "").trim();

  if (normalizedUser.length < 3) {
    return { ok: false, message: "O usuario precisa ter pelo menos 3 caracteres." };
  }

  if (normalizedPassword.length < 4) {
    return { ok: false, message: "A senha precisa ter pelo menos 4 caracteres." };
  }

  if (findUser(normalizedUser)) {
    return { ok: false, message: "Esse usuario ja existe." };
  }

  const users = readUsers();
  const account = {
    userName: normalizedUser,
    password: normalizedPassword,
    createdAt: nowIso()
  };
  users.push(account);
  writeUsers(users);
  rememberHint(normalizedUser);
  return { ok: true, account };
}

function loginWithAccount(userName, password) {
  const account = findUser(userName);
  if (!account) {
    return { ok: false, message: "Conta nao encontrada. Crie uma conta antes de entrar." };
  }

  if (String(account.password || "") !== String(password || "").trim()) {
    return { ok: false, message: "Senha invalida para este usuario." };
  }

  const session = writeSession(account.userName);
  return { ok: true, session, account };
}

window.PCPAuth = {
  readSession,
  readUsers,
  writeSession,
  clearSession,
  touchSession,
  isSessionValid,
  formatDateTime,
  getRememberedHint,
  createAccount,
  loginWithAccount,
  findUser,
  keys: {
    session: PCP_AUTH_KEY,
    hint: PCP_AUTH_HINT_KEY,
    users: PCP_AUTH_USERS_KEY
  }
};
