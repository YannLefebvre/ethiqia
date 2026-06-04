import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://svvcfsnjkmirlxqvvxmw.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN2dmNmc25qa21pcmx4cXZ2eG13Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA1NDAyNzMsImV4cCI6MjA5NjExNjI3M30.FPZsc5i4Mhkz7XCN7xrUU-cW0DWMDNSWHOsFf8SrQ5Q";
const HEADERS = { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${SUPABASE_KEY}`, "Prefer": "resolution=merge-duplicates", "X-Upsert": "true" };

const questions = [
  { id: 1, titre: "IA mentor", altA: "Oui, c'est plus facile et plus accessible qu'un prof qui a rarement le temps de me guider.", altB: "Non, même si j'ai besoin d'un coach, je préfère tracer mon chemin par mes propres moyens." },
  { id: 2, titre: "IA tuteur", altA: "Pourquoi pas, le prof n'est pas toujours disponible pour m'aider.", altB: "Surtout pas, ça reste une machine, elle ne me connaît pas et elle demande trop de données personnelles." },
  { id: 3, titre: "IA équipière", altA: "Oui, ce serait pas mal d'avoir un autre avis, même si c'est celui d'une machine.", altB: "Non, c'est déjà difficile de s'entendre entre nous, alors avec une voix en plus…" },
  { id: 4, titre: "IA coach", altA: "Oui, j'ai toujours rêvé d'avoir une voix qui m'aiderait à savoir où j'en suis !", altB: "Non, ça va pas du tout ! Elle va collecter plein de données sur mon travail. C'est trop personnel pour lui laisser tout ça !" },
  { id: 5, titre: "IA simulateur", altA: "Oui, c'est quand même mieux qu'une étude de cas avec plein de questions.", altB: "Non, une simulation reste une simulation : elle ne représente pas assez bien ce qu'on voit dans la vraie vie." },
  { id: 6, titre: "IA outil", altA: "Oui, elle sera plus efficace sur ce qu'on cherche à faire.", altB: "Non, on a besoin de travailler ces compétences. Avec l'IA, on ne pourra jamais les entraîner correctement." },
  { id: 7, titre: "IA soutien", altA: "Oui, sans IA, il est bien difficile de leur apporter de l'aide.", altB: "Non, ce n'est pas juste pour les autres qui ne bénéficieront pas de cette IA." },
  { id: 8, titre: "Recherche documentaire", altA: "Une IA générative : on passe plus rapidement au travail de compréhension, de liens, de traitement des données.", altB: "Un moteur de recherche : ça demande de critiquer les résultats et de sélectionner les informations utiles, avec leurs sources." },
  { id: 9, titre: "Un résultat préoccupant", altA: "Je retente, jusqu'à ce que le résultat soit satisfaisant.", altB: "Je retente mais pas trop, il paraît que ça consomme beaucoup d'énergie et de ressources." },
  { id: 10, titre: "Justice ?", altA: "Je réagis. Il faut que les profs soient capables de gérer cette injustice.", altB: "Je ne fais rien. Ça se reproduira, et ça ne va pas m'encourager à continuer comme ça." },
  { id: 11, titre: "Confiance ?", altA: "Je continuerai à faire comme ça. Je n'ai rien à me reprocher.", altB: "J'utilise un site qui permet de vérifier si mon texte est ressemblant à ce qu'une IA pourrait produire." },
  { id: 12, titre: "Gaspillage", altA: "Tant pis, si je me préoccupe de cela, je ne ferai plus grand-chose.", altB: "Quand même, ça me gêne. Je vais essayer de trouver des pratiques un peu plus responsables." },
  { id: 13, titre: "Grand ou petit modèle ?", altA: "Oui, j'ai l'habitude avec ces outils. C'est plus simple.", altB: "Non, il paraît que les modèles en local sont plus petits, moins consommateurs, mais ils sont moins accessibles." },
  { id: 14, titre: "Où se trouve la source ?", altA: "Tant pis, on fera sans, ce n'est pas si important.", altB: "Quand même, je vais utiliser un moteur de recherche ou des livres, ça existe encore." },
  { id: 15, titre: "Les deux à la fois ?", altA: "C'est parfait, je ne suis plus obligé de cliquer sur les liens et de lire les résultats !", altB: "C'est limité, je ne peux même plus faire mon choix d'informations moi-même !" },
  { id: 16, titre: "Aide à l'écriture", altA: "Quitte à faire, tout le travail à ma place.", altB: "De la reformulation du texte que j'ai écrit moi-même." },
  { id: 17, titre: "Transparence", altA: "Je rends mon travail, sans aucune précision.", altB: "Je copie et colle ma demande à l'IA, le prof saura comment je l'ai utilisée." },
  { id: 18, titre: "Nourrir une IA", altA: "Aucun problème, même si l'IA va conserver mon fichier et le réutiliser plus tard.", altB: "Pas question de la laisser utiliser ce que j'ai fait. C'est moi le propriétaire." },
  { id: 19, titre: "Calculatrice de luxe", altA: "Oui, ça permet d'aller plus vite.", altB: "Non, inutile d'utiliser un outil aussi puissant et consommateur pour un truc aussi facile." },
  { id: 20, titre: "Qui fera le résumé ?", altA: "Je demande à l'IA de faire la synthèse de ce que j'ai trouvé. Rapide et efficace.", altB: "Je fais la synthèse moi-même, sinon, je ne comprendrai rien au sujet." },
  { id: 21, titre: "Ultra rapide", altA: "Je fais la demande comme sur un moteur de recherche, quitte à l'améliorer par la suite.", altB: "Je prends le temps de bien faire ma première demande. Ça m'évitera d'en faire d'autres pour préciser la réponse." },
  { id: 22, titre: "Doutes", altA: "Je compare avec les résultats d'autres IA.", altB: "Je reste sur ces résultats. D'autres IA ne m'apporteront rien de plus." },
  { id: 23, titre: "En une fois ou en plein de fois", altA: "Je le fais en plusieurs fois, en précisant mes demandes.", altB: "Je travaille ma demande pour avoir un résultat acceptable rapidement." },
  { id: 24, titre: "Théorie du remplacement", altA: "Même pas peur. On trouvera d'autres emplois dans le futur, on s'adaptera.", altB: "Ça me stresse. Je ne vais pas commencer à me former à un métier qui pourrait disparaître." },
  { id: 25, titre: "Un peu d'entraînement", altA: "Je lui donne uniquement les cours de mes profs.", altB: "Je lui demande de chercher où elle veut, sans la restreindre." },
  { id: 26, titre: "Intuition de départ", altA: "Je demande à l'IA de me donner le départ. Je continuerai moi-même ensuite.", altB: "Je veux commencer moi-même, sinon, je ne maîtrise pas le sujet. L'IA pourra m'aider plus tard, pour d'autres tâches." },
  { id: 27, titre: "Absence de cadre", altA: "Je préfère éviter au maximum de l'utiliser, ça pourrait toujours m'être reproché.", altB: "Je l'utilise quand même, s'il n'y a pas de cadre, c'est qu'il n'y a pas de problème." },
  { id: 28, titre: "Coaching par l'IA", altA: "Oui, ça peut toujours m'aider.", altB: "Non, il lui faudra plein de données sur moi, et je n'ai pas envie de ça." },
  { id: 29, titre: "L'IA va m'orienter", altA: "Oui, ça ne sera pas pire que les jurys qui ne sont pas toujours justes.", altB: "Non, c'est mon avenir et mes données, je ne veux pas qu'une IA puisse influencer mon destin." },
  { id: 30, titre: "Une IA magique", altA: "Oui, c'est tellement plus rapide que de décomposer toutes ces tâches !", altB: "Non, je n'ai aucun moyen de vérifier et de valider ce qu'elle fait, et de corriger s'il y a besoin." },
  { id: 31, titre: "Chatbot motivationnel", altA: "Oui, ça aide quand on ne trouve personne pour se rebooster.", altB: "Non, ça reste une machine programmée, pas un ami ou un psy." },
  { id: 32, titre: "IA sélective 1", altA: "On les réserve pour les personnes les plus bosseuses, pour qu'elles continuent à performer.", altB: "On les réserve aux personnes en difficulté, pour les aider à s'améliorer." },
  { id: 33, titre: "IA sélective 2", altA: "On les réserve pour une partie de la classe, tant pis pour l'égalité.", altB: "Soit c'est l'IA pour tout le monde, soit c'est l'IA pour personne." },
  { id: 34, titre: "Trop de triche !", altA: "Ils devraient tout interdire et renforcer les moyens pour le contrôler.", altB: "Ils devraient nous encourager à dire quand et comment on utilise l'IA, notamment pour les DM." },
  { id: 35, titre: "Illustrations", altA: "Je demande à une IA générative pour les créer rapidement, selon mes instructions.", altB: "Je demande à un moteur de recherche, surtout que je ne sais pas exactement ce que je voudrais." },
  { id: 36, titre: "Exposé", altA: "Je préfère l'IA, c'est plus rapide et pratique.", altB: "Je préfère le groupe de travail, c'est plus convivial." },
  { id: 37, titre: "Supériorité", altA: "C'est comme ça, c'est possible que ce soit ces personnes qui ont marqué l'histoire des sciences.", altB: "Il doit y avoir un biais, puisque c'est une IA américaine." },
  { id: 38, titre: "Mettre les mains dans le cambouis", altA: "Inutile. Je ne veux pas devenir informaticien !", altB: "Utile, j'ai besoin de savoir comment ces outils marchent pour mieux comprendre leurs résultats." },
  { id: 39, titre: "Bleu, blanc, rouge ?", altA: "Je préfère le modèle français, même s'il était un peu plus cher.", altB: "Je préfère le modèle américain ou chinois, peut-être plus performant et plus utilisé dans le monde." },
];

const emptyCount = () => ({ A: 0, B: 0 });

async function fetchCounts() {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/votes?select=question_id,option,count`, {
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${SUPABASE_KEY}` }
  });
  if (!res.ok) throw new Error("Erreur de chargement");
  const rows = await res.json();
  const counts = {};
  for (const row of rows) {
    if (!counts[row.question_id]) counts[row.question_id] = emptyCount();
    counts[row.question_id][row.option] = row.count;
  }
  return counts;
}

async function upsertCount(question_id, option, count) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/votes?on_conflict=question_id,option`, {
    method: "POST",
    headers: HEADERS,
    body: JSON.stringify({ question_id, option, count })
  });
  if (!res.ok) throw new Error("Erreur de sauvegarde");
}

export default function App() {
  const [counts, setCounts] = useState({});
  const [view, setView] = useState("form");
  const [status, setStatus] = useState("loading"); // loading | ready | error | saving
  const [errorMsg, setErrorMsg] = useState("");

  const load = useCallback(async () => {
    setStatus("loading");
    try {
      const data = await fetchCounts();
      setCounts(data);
      setStatus("ready");
    } catch (e) {
      setStatus("error");
      setErrorMsg(e.message);
    }
  }, []);

  useEffect(() => { load(); }, [load]);

  const handleChoice = async (id, option) => {
    const prev = counts[id] || emptyCount();
    const newCount = prev[option] + 1;
    const optimistic = { ...counts, [id]: { ...prev, [option]: newCount } };
    setCounts(optimistic);
    setStatus("saving");
    try {
      await upsertCount(id, option, newCount);
      setStatus("ready");
    } catch {
      setCounts(counts); // rollback
      setStatus("error");
      setErrorMsg("Impossible de sauvegarder. Vérifiez votre connexion.");
    }
  };

  const decrement = async (id, option) => {
    const prev = counts[id] || emptyCount();
    if (prev[option] <= 0) return;
    const newCount = prev[option] - 1;
    const optimistic = { ...counts, [id]: { ...prev, [option]: newCount } };
    setCounts(optimistic);
    setStatus("saving");
    try {
      await upsertCount(id, option, newCount);
      setStatus("ready");
    } catch {
      setCounts(counts);
      setStatus("error");
      setErrorMsg("Impossible de sauvegarder.");
    }
  };

  const clearAll = async () => {
    if (!window.confirm("Remettre tous les compteurs à zéro dans Supabase ?")) return;
    setStatus("saving");
    try {
      // Set all to 0 via upserts
      await Promise.all(
        questions.flatMap(q => [
          upsertCount(q.id, "A", 0),
          upsertCount(q.id, "B", 0),
        ])
      );
      setCounts({});
      setStatus("ready");
    } catch {
      setStatus("error");
      setErrorMsg("Réinitialisation échouée.");
    }
  };

  const exportXLSX = () => {
    const rows = questions.map(q => {
      const c = counts[q.id] || emptyCount();
      const tot = c.A + c.B;
      return {
        "N°": q.id,
        "Titre": q.titre,
        "Option 1 (texte)": q.altA,
        "Option 2 (texte)": q.altB,
        "Votes Option 1": c.A,
        "Votes Option 2": c.B,
        "Total votes": tot,
        "% Option 1": tot > 0 ? Math.round((c.A / tot) * 100) + "%" : "-",
        "% Option 2": tot > 0 ? Math.round((c.B / tot) * 100) + "%" : "-",
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 4 }, { wch: 28 }, { wch: 60 }, { wch: 60 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
    const totalA = Object.values(counts).reduce((s, c) => s + (c.A || 0), 0);
    const totalB = Object.values(counts).reduce((s, c) => s + (c.B || 0), 0);
    const ws2 = XLSX.utils.json_to_sheet([
      { "": "Total votes Option 1", "Valeur": totalA },
      { "": "Total votes Option 2", "Valeur": totalB },
      { "": "Total général", "Valeur": totalA + totalB },
    ]);
    ws2["!cols"] = [{ wch: 24 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Résultats");
    XLSX.utils.book_append_sheet(wb, ws2, "Synthèse");
    XLSX.writeFile(wb, "ethiqia_résultats.xlsx");
  };

  const totalVotes = Object.values(counts).reduce((s, c) => s + (c.A || 0) + (c.B || 0), 0);
  const totalA = Object.values(counts).reduce((s, c) => s + (c.A || 0), 0);
  const totalB = Object.values(counts).reduce((s, c) => s + (c.B || 0), 0);

  const statusDot = { loading: "#ffd764", saving: "#4fc3f7", ready: "#69f0ae", error: "#ff5252" };
  const statusLabel = { loading: "Chargement…", saving: "Sauvegarde…", ready: "Synchronisé ✓", error: "Erreur" };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%)", fontFamily: "'Georgia', serif", color: "#e8e0d0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=Source+Serif+4:ital,wght@0,300;0,400;1,300&display=swap');
        * { box-sizing: border-box; }
        .card { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.12); border-radius: 16px; padding: 22px; margin-bottom: 18px; transition: all 0.3s ease; }
        .card:hover { background: rgba(255,255,255,0.07); }
        .vote-row { display: flex; align-items: center; gap: 10px; }
        .btn-vote { flex: 1; text-align: left; padding: 13px 16px; border-radius: 12px; border: 2px solid rgba(255,255,255,0.1); cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 13.5px; line-height: 1.5; transition: all 0.2s; background: rgba(255,255,255,0.05); color: #e8e0d0; }
        .btn-vote.opt-a:hover { border-color: #4fc3f7; background: rgba(79,195,247,0.1); }
        .btn-vote.opt-b:hover { border-color: #ce93d8; background: rgba(206,147,216,0.1); }
        .btn-vote.opt-a.has-votes { border-color: rgba(79,195,247,0.5); background: rgba(79,195,247,0.1); }
        .btn-vote.opt-b.has-votes { border-color: rgba(206,147,216,0.5); background: rgba(206,147,216,0.1); }
        .count-badge { min-width: 36px; height: 36px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: bold; flex-shrink: 0; transition: all 0.15s; border: none; }
        .count-a { background: rgba(79,195,247,0.25); color: #4fc3f7; cursor: pointer; }
        .count-a:hover { background: rgba(79,195,247,0.4); }
        .count-b { background: rgba(206,147,216,0.25); color: #ce93d8; cursor: pointer; }
        .count-b:hover { background: rgba(206,147,216,0.4); }
        .count-zero { background: rgba(255,255,255,0.06); color: #665e52; cursor: default; }
        .tab { padding: 8px 18px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.2); background: transparent; color: #b0a898; cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 13px; transition: all 0.2s; }
        .tab.active { background: rgba(255,215,100,0.15); border-color: rgba(255,215,100,0.5); color: #ffd764; }
        .btn-action { padding: 9px 20px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.2); background: rgba(255,255,255,0.07); color: #e8e0d0; cursor: pointer; font-family: 'Source Serif 4', serif; font-size: 13px; transition: all 0.2s; }
        .btn-action:hover { background: rgba(255,255,255,0.14); }
        .summary-row { display: flex; align-items: center; gap: 12px; padding: 11px 0; border-bottom: 1px solid rgba(255,255,255,0.06); }
        .spinner { width: 16px; height: 16px; border: 2px solid rgba(255,255,255,0.1); border-top-color: #4fc3f7; border-radius: 50%; animation: spin 0.7s linear infinite; display: inline-block; }
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>

      <div style={{ maxWidth: 780, margin: "0 auto", padding: "32px 16px" }}>

        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: "clamp(28px,5vw,46px)", fontWeight: 900, margin: 0, background: "linear-gradient(135deg, #ffd764, #ce93d8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Éthiq·IA</h1>
          <p style={{ fontFamily: "'Source Serif 4', serif", fontStyle: "italic", color: "#a09888", fontSize: 14, marginTop: 5 }}>Cliquez sur une option pour voter — votes sauvegardés en ligne</p>
        </div>

        {/* Status + stats */}
        <div style={{ marginBottom: 22, padding: "14px 20px", background: "rgba(255,255,255,0.05)", borderRadius: 12, border: "1px solid rgba(255,255,255,0.1)", display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
          <div><span style={{ fontSize: 12, color: "#a09888" }}>Total votes</span><p style={{ margin: "2px 0 0", fontSize: 22, fontWeight: "bold", color: "#ffd764" }}>{totalVotes}</p></div>
          <div><span style={{ fontSize: 12, color: "#a09888" }}>Option 1</span><p style={{ margin: "2px 0 0", fontSize: 22, fontWeight: "bold", color: "#4fc3f7" }}>{totalA}</p></div>
          <div><span style={{ fontSize: 12, color: "#a09888" }}>Option 2</span><p style={{ margin: "2px 0 0", fontSize: 22, fontWeight: "bold", color: "#ce93d8" }}>{totalB}</p></div>
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
            {status === "loading" || status === "saving" ? <span className="spinner" /> : null}
            <span style={{ width: 8, height: 8, borderRadius: "50%", background: statusDot[status], display: "inline-block" }} />
            <span style={{ fontSize: 12, color: status === "error" ? "#ff5252" : "#a09888" }}>{statusLabel[status]}</span>
            {status === "error" && <button className="btn-action" style={{ padding: "4px 10px", fontSize: 11 }} onClick={load}>Réessayer</button>}
          </div>
        </div>

        {status === "error" && errorMsg && (
          <div style={{ background: "rgba(255,82,82,0.1)", border: "1px solid rgba(255,82,82,0.3)", borderRadius: 10, padding: "12px 16px", marginBottom: 18, fontSize: 13, color: "#ff8a80" }}>
            ⚠️ {errorMsg}
          </div>
        )}

        {/* Nav */}
        <div style={{ display: "flex", gap: 8, marginBottom: 22, flexWrap: "wrap", alignItems: "center" }}>
          <button className={`tab ${view === "form" ? "active" : ""}`} onClick={() => setView("form")}>📋 Questionnaire</button>
          <button className={`tab ${view === "summary" ? "active" : ""}`} onClick={() => setView("summary")}>📊 Récapitulatif</button>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <button className="btn-action" onClick={load} title="Rafraîchir depuis Supabase">↺ Actualiser</button>
            <button className="btn-action" style={{ background: "rgba(79,195,247,0.15)", borderColor: "rgba(79,195,247,0.4)", color: "#4fc3f7" }} onClick={exportXLSX}>⬇ Exporter .xlsx</button>
            <button className="btn-action" style={{ color: "#ff8a80" }} onClick={clearAll}>✕ Réinitialiser</button>
          </div>
        </div>

        {status === "loading" && (
          <div style={{ textAlign: "center", padding: "60px 0", color: "#a09888" }}>
            <span className="spinner" style={{ width: 28, height: 28, borderWidth: 3 }} />
            <p style={{ marginTop: 16 }}>Connexion à Supabase…</p>
          </div>
        )}

        {status !== "loading" && view === "form" && questions.map(q => {
          const c = counts[q.id] || emptyCount();
          const tot = c.A + c.B;
          const pctA = tot > 0 ? (c.A / tot) * 100 : 50;
          return (
            <div key={q.id} className="card">
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <div style={{ width: 26, height: 26, borderRadius: "50%", background: "rgba(255,255,255,0.07)", color: "#a09888", fontSize: 11, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{q.id}</div>
                <h3 style={{ margin: 0, fontFamily: "'Playfair Display', serif", fontSize: 16, fontWeight: 700, color: "#f5efe0" }}>{q.titre}</h3>
                {tot > 0 && <span style={{ marginLeft: "auto", fontSize: 11, color: "#665e52" }}>{tot} vote{tot > 1 ? "s" : ""}</span>}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div className="vote-row">
                  <button className={`btn-vote opt-a${c.A > 0 ? " has-votes" : ""}`} onClick={() => handleChoice(q.id, "A")}>
                    <span style={{ fontSize: 10, fontWeight: "bold", color: "#4fc3f7", display: "block", marginBottom: 3 }}>OPTION 1 — cliquer pour voter</span>
                    {q.altA}
                  </button>
                  <button className={`count-badge ${c.A > 0 ? "count-a" : "count-zero"}`} title={c.A > 0 ? "Retirer un vote" : ""} onClick={() => decrement(q.id, "A")}>{c.A}</button>
                </div>
                <div className="vote-row">
                  <button className={`btn-vote opt-b${c.B > 0 ? " has-votes" : ""}`} onClick={() => handleChoice(q.id, "B")}>
                    <span style={{ fontSize: 10, fontWeight: "bold", color: "#ce93d8", display: "block", marginBottom: 3 }}>OPTION 2 — cliquer pour voter</span>
                    {q.altB}
                  </button>
                  <button className={`count-badge ${c.B > 0 ? "count-b" : "count-zero"}`} title={c.B > 0 ? "Retirer un vote" : ""} onClick={() => decrement(q.id, "B")}>{c.B}</button>
                </div>
                {tot > 0 && (
                  <div style={{ height: 4, borderRadius: 2, background: "rgba(255,255,255,0.08)", overflow: "hidden" }}>
                    <span style={{ width: `${pctA}%`, height: "100%", background: "#4fc3f7", display: "inline-block" }} />
                    <span style={{ width: `${100 - pctA}%`, height: "100%", background: "#ce93d8", display: "inline-block" }} />
                  </div>
                )}
              </div>
            </div>
          );
        })}

        {status !== "loading" && view === "summary" && (
          <div style={{ background: "rgba(255,255,255,0.04)", borderRadius: 16, border: "1px solid rgba(255,255,255,0.1)", padding: "20px 24px" }}>
            {questions.map(q => {
              const c = counts[q.id] || emptyCount();
              const tot = c.A + c.B;
              const pctA = tot > 0 ? Math.round((c.A / tot) * 100) : null;
              return (
                <div key={q.id} className="summary-row">
                  <div style={{ width: 24, height: 24, borderRadius: "50%", background: "rgba(255,255,255,0.06)", color: "#665e52", fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{q.id}</div>
                  <div style={{ flex: 1 }}>
                    <p style={{ margin: "0 0 4px", fontFamily: "'Playfair Display', serif", fontSize: 13.5, color: "#f5efe0" }}>{q.titre}</p>
                    {tot > 0 ? (
                      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                        <span style={{ fontSize: 12, color: "#4fc3f7" }}>Opt. 1 : <b>{c.A}</b></span>
                        <span style={{ fontSize: 12, color: "#ce93d8" }}>Opt. 2 : <b>{c.B}</b></span>
                        <span style={{ fontSize: 11, color: "#665e52" }}>({tot} vote{tot > 1 ? "s" : ""} — {pctA}% / {100 - pctA}%)</span>
                        <div style={{ flex: 1, minWidth: 80, height: 4, background: "rgba(255,255,255,0.08)", borderRadius: 2, overflow: "hidden" }}>
                          <div style={{ width: `${pctA}%`, height: "100%", background: "#4fc3f7", display: "inline-block" }} />
                          <div style={{ width: `${100 - pctA}%`, height: "100%", background: "#ce93d8", display: "inline-block" }} />
                        </div>
                      </div>
                    ) : (
                      <span style={{ fontSize: 11, color: "#443d36", fontStyle: "italic" }}>Aucun vote</span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        <p style={{ textAlign: "center", fontSize: 11, color: "#443d36", marginTop: 28 }}>
          Votes sauvegardés dans Supabase · Cliquer sur le compteur coloré pour retirer un vote
        </p>
      </div>
    </div>
  );
}
