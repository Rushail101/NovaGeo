// Sign-in, roles and the gate in front of the whole finance app.
// The database (supabase/001_security.sql) is the real lock; this file decides what to show.
import React, { createContext, useContext, useEffect, useState } from "react";
import { supabase } from "./supabase.js";

export const ROLE_LABELS = {
  admin: "Admin / Partner",
  accountant: "Accountant",
  viewer: "Viewer (read only)",
  pending: "Pending approval",
};

const AuthCtx = createContext({ session: null, role: "", name: "", email: "" });
export const useAuth = () => useContext(AuthCtx);
export const canWrite = role => role === "admin" || role === "accountant";

async function loadProfile(userId) {
  const { data, error } = await supabase.from("fin_users").select("role, name, email, active").eq("user_id", userId).maybeSingle();
  if (error) return { setupNeeded: /fin_users/.test(error.message || ""), role: "", error: error.message };
  if (!data) return { role: "pending" };
  return { role: data.active ? data.role : "blocked", name: data.name, email: data.email };
}

export function AuthGate({ children }) {
  const [session, setSession] = useState(undefined);
  const [profile, setProfile] = useState(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setProfile(null); return; }
    let live = true;
    loadProfile(session.user.id).then(p => live && setProfile(p));
    return () => { live = false; };
  }, [session?.user?.id]);

  if (session === undefined || (session && !profile)) return <Splash text="Checking sign-in…" />;
  if (!session) return <Login />;
  if (profile.setupNeeded) return <Notice title="Security setup not run yet" signOut>
    Run <span className="mono">supabase/001_security.sql</span> in the Supabase SQL editor, then reload this page.
    Until then the database is open to anyone who has the site address.
  </Notice>;
  if (profile.role === "blocked") return <Notice title="Account blocked" signOut>Ask an admin to unblock your account.</Notice>;
  if (!["admin", "accountant", "viewer"].includes(profile.role)) return <Notice title="Waiting for approval" signOut>
    Your account <b>{session.user.email}</b> is created. An admin needs to approve it under <b>Admin › Users</b> before you can see the books.
    <div style={{ marginTop: 12 }}><button className="btn btn-ghost btn-sm" onClick={() => location.reload()}>Check again</button></div>
  </Notice>;

  const value = { session, role: profile.role, name: profile.name || session.user.email, email: session.user.email };
  return <AuthCtx.Provider value={value}>{children}</AuthCtx.Provider>;
}

export const signOut = () => supabase.auth.signOut();

function Shell({ children }) {
  return (
    <div className="auth-wrap">
      <div className="auth-card">
        <div className="auth-brand"><span className="auth-mark">NG</span><div><b>Finance ERP</b><small>Nova Geo Resources LLP</small></div></div>
        {children}
      </div>
    </div>
  );
}

function Splash({ text }) {
  return <div className="auth-wrap"><div className="mono" style={{ color: "var(--accent)" }}>{text}</div></div>;
}

function Notice({ title, children, signOut: showSignOut }) {
  return (
    <Shell>
      <h3 style={{ fontSize: 15, marginBottom: 8 }}>{title}</h3>
      <div style={{ color: "var(--text2)", fontSize: 13, lineHeight: 1.6 }}>{children}</div>
      {showSignOut && <button className="btn btn-ghost" style={{ marginTop: 18 }} onClick={signOut}>Sign out</button>}
    </Shell>
  );
}

function Login() {
  const [mode, setMode] = useState("in"); // in | up | reset
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setMsg(null);
    try {
      if (mode === "in") {
        const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (error) throw new Error("Wrong email or password.");
      } else if (mode === "up") {
        if (password.length < 10) throw new Error("Use at least 10 characters for the password.");
        const { data, error } = await supabase.auth.signUp({ email: email.trim(), password, options: { data: { name: name.trim() } } });
        if (error) throw error;
        if (!data.session) setMsg({ ok: true, text: "Account created. Check your email to confirm it, then sign in." });
      } else {
        const { error } = await supabase.auth.resetPasswordForEmail(email.trim(), { redirectTo: location.origin });
        if (error) throw error;
        setMsg({ ok: true, text: "If that email has an account, a reset link is on its way." });
      }
    } catch (err) {
      setMsg({ ok: false, text: err.message || String(err) });
    } finally {
      setBusy(false);
    }
  }

  return (
    <Shell>
      <form onSubmit={submit} className="auth-form">
        {mode === "up" && <div className="form-group"><label>Your name</label><input value={name} onChange={e => setName(e.target.value)} required /></div>}
        <div className="form-group"><label>Email</label><input type="email" autoComplete="username" value={email} onChange={e => setEmail(e.target.value)} required /></div>
        {mode !== "reset" && <div className="form-group"><label>Password</label><input type="password" autoComplete={mode === "up" ? "new-password" : "current-password"} value={password} onChange={e => setPassword(e.target.value)} required /></div>}
        {msg && <div className={msg.ok ? "auth-ok" : "auth-err"}>{msg.text}</div>}
        <button className="btn btn-primary" style={{ justifyContent: "center" }} disabled={busy}>
          {busy ? "Please wait…" : mode === "in" ? "Sign in" : mode === "up" ? "Create account" : "Send reset link"}
        </button>
        <div className="auth-links">
          {mode !== "in" && <button type="button" onClick={() => { setMode("in"); setMsg(null); }}>Back to sign in</button>}
          {mode === "in" && <button type="button" onClick={() => { setMode("up"); setMsg(null); }}>New user? Create account</button>}
          {mode === "in" && <button type="button" onClick={() => { setMode("reset"); setMsg(null); }}>Forgot password</button>}
        </div>
      </form>
    </Shell>
  );
}

// Shown after a password-reset email link: lets the user set a new password.
export function PasswordRecovery() {
  const [show, setShow] = useState(false);
  const [pw, setPw] = useState("");
  const [msg, setMsg] = useState("");
  useEffect(() => {
    const { data: sub } = supabase.auth.onAuthStateChange(e => { if (e === "PASSWORD_RECOVERY") setShow(true); });
    return () => sub.subscription.unsubscribe();
  }, []);
  if (!show) return null;
  async function save() {
    if (pw.length < 10) return setMsg("Use at least 10 characters.");
    const { error } = await supabase.auth.updateUser({ password: pw });
    if (error) return setMsg(error.message);
    setShow(false);
  }
  return (
    <div className="modal-overlay"><div className="modal" style={{ maxWidth: 400 }}>
      <div className="modal-head"><h3>Set a new password</h3></div>
      <div className="modal-body">
        <div className="form-group"><label>New password</label><input type="password" value={pw} onChange={e => setPw(e.target.value)} autoComplete="new-password" /></div>
        {msg && <div className="auth-err" style={{ marginTop: 10 }}>{msg}</div>}
      </div>
      <div className="modal-foot"><button className="btn btn-primary" onClick={save}>Save password</button></div>
    </div></div>
  );
}

export const AUTH_CSS = `
.auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;background:radial-gradient(ellipse at top,var(--bg3),var(--bg) 60%)}
.auth-card{width:100%;max-width:380px;background:var(--bg2);border:1px solid var(--border2);border-radius:var(--r3);padding:26px;box-shadow:var(--shadow2)}
.auth-brand{display:flex;gap:12px;align-items:center;margin-bottom:22px}
.auth-brand b{display:block;font-size:17px;color:var(--accent)}
.auth-brand small{font-family:var(--mono);font-size:10px;color:var(--text3);letter-spacing:.08em;text-transform:uppercase}
.auth-mark{width:38px;height:38px;border-radius:10px;background:var(--accent);color:#0d0d0d;display:grid;place-items:center;font-weight:800}
.auth-form{display:flex;flex-direction:column;gap:14px}
.auth-err{color:var(--red);font-size:12.5px;background:#200808;border:1px solid #401010;border-radius:var(--r);padding:8px 10px}
.auth-ok{color:var(--green);font-size:12.5px;background:#0a2016;border:1px solid #1a4028;border-radius:var(--r);padding:8px 10px}
.auth-links{display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap}
.auth-links button{background:none;border:none;color:var(--text2);font-size:12px;cursor:pointer;font-family:var(--font);padding:0}
.auth-links button:hover{color:var(--accent)}
`;
