import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Plus, X, Check, ChevronDown, ChevronRight, ChevronLeft, ChevronUp,
  Edit2, Trash2, Search, Copy, MoreVertical, Settings, Download, Upload,
  TrendingUp, ListChecks, FileText, AlertCircle, Wallet, Eye, EyeOff,
  ArrowRight, RotateCcw, Filter, CheckCircle2, Calculator, Share2,
  Zap, Hash, Users, Banknote, Sparkles, Info, Pencil, ArrowUpDown,
  Send, RefreshCw, Database, ScrollText, Loader2, ArrowLeftRight,
} from "lucide-react";

// ═══════════════════════════════════════════════════════════════════
//  THEME · NEBULOSA
// ═══════════════════════════════════════════════════════════════════
const C = {
  bg: "#0a0612",
  surface: "rgba(16, 10, 28, 0.88)",
  card: "rgba(28, 18, 42, 0.72)",
  cardHover: "rgba(36, 24, 56, 0.88)",
  border: "rgba(180, 100, 220, 0.32)",
  borderStrong: "rgba(180, 100, 220, 0.5)",
  accent: "#e84cc1",
  accent2: "#22d3ee",
  accentDim: "rgba(232, 76, 193, 0.14)",
  accent2Dim: "rgba(34, 211, 238, 0.14)",
  text: "#fefdff",
  textDim: "#f2e5f8",
  soft: "#cebede",
  muted: "#a595bb",
  nav: "rgba(8, 4, 16, 0.96)",
  warn: "#fbbf24",
  warnDim: "rgba(251, 191, 36, 0.14)",
  positive: "#5eead4",
  positiveDim: "rgba(94, 234, 212, 0.12)",
  positiveBorder: "rgba(94, 234, 212, 0.35)",
  negative: "#fb7185",
  negativeDim: "rgba(251, 113, 133, 0.12)",
  negativeBorder: "rgba(251, 113, 133, 0.35)",
};
const T = {
  compra: { primary: "#5eead4", dim: "rgba(94,234,212,0.07)", border: "rgba(94,234,212,0.24)", bg: "rgba(94,234,212,0.1)", glow: "0 0 24px rgba(94,234,212,0.14)", label: "COMPRA" },
  venta: { primary: "#fb7185", dim: "rgba(251,113,133,0.07)", border: "rgba(251,113,133,0.24)", bg: "rgba(251,113,133,0.1)", glow: "0 0 24px rgba(251,113,133,0.14)", label: "VENTA" },
};
const EASE = "cubic-bezier(0.16, 1, 0.3, 1)";
const NEAR_COMPLETE = 50000;

// ═══════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════
const uid = () => Math.random().toString(36).slice(2, 11);
const todayStr = () => { const d = new Date(); return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`; };
const fmtARS = n => "$" + Math.round(n ?? 0).toLocaleString("es-AR");
const fmtARSdec = n => "$" + (n ?? 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtUSDT = n => (n ?? 0).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " USDT";
const fmtPct = n => (n ?? 0).toFixed(2).replace(".", ",") + "%";
const parseNum = s => { const v = String(s || "0").replace(/\s/g, "").replace(/\./g, "").replace(",", "."); return parseFloat(v) || 0; };
const MONTHS = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const DAYS = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];
const monthLabel = today => { const [, mm, yyyy] = today.split("/"); return `${MONTHS[+mm - 1]} ${yyyy}`; };
const dayName = dateStr => { const [dd, mm, yyyy] = dateStr.split("/"); return DAYS[new Date(+yyyy, +mm - 1, +dd).getDay()]; };
const dateToTime = ds => { const [dd, mm, yyyy] = ds.split("/"); return new Date(+yyyy, +mm - 1, +dd).getTime(); };
const dateMonthKey = ds => ds.slice(3);

// Available months from ops
const availableMonths = ops => {
  const set = new Set(ops.map(o => o.date.slice(3)));
  return [...set].sort((a, b) => {
    const [ma, ya] = a.split("/"), [mb, yb] = b.split("/");
    return new Date(+yb, +mb - 1) - new Date(+ya, +ma - 1);
  });
};

// ═══════════════════════════════════════════════════════════════════
//  BANK CONFIG
// ═══════════════════════════════════════════════════════════════════
const BANKS = [
  { id: "Ripio", comm: 0, label: "Ripio" },
  { id: "Lemon", comm: 0, label: "Lemon" },
  { id: "Fiwind", comm: 0, label: "Fiwind" },
  { id: "Letsbit", comm: 0.0035, label: "Letsbit (0,35% in)" },
  { id: "Telepagos", comm: 0.0045, label: "Telepagos (0,45% in)" },
  { id: "Cocos", comm: 0, label: "Cocos" },
  { id: "Mercadopago", comm: 0, label: "Mercadopago" },
  { id: "Galicia", comm: 0, label: "Galicia" },
];
const BINANCE_FEE = 0.0016;

// ═══════════════════════════════════════════════════════════════════
//  INITIAL DATA (historial mayo preservado)
// ═══════════════════════════════════════════════════════════════════
const INITIAL_DATA = [{"id":"nimnpu06n","date":"01/05/2026","client":"VERO","type":"compra","usdtAmount":3574.64,"tcClient":1467.0,"tcOtc":null,"bank":"Ripio","exchange":"","tts":[{"id":"dz062s6zx","amount":5244000.0,"ts":1778714762508}],"ts":1778714762508},{"id":"ife2ms5se","date":"04/05/2026","client":"PABLITO","type":"venta","usdtAmount":14247.31,"tcClient":1488.0,"tcOtc":null,"bank":"Ripio","exchange":"Binance","tts":[{"id":"zee3h3pqs","amount":21125800.0,"ts":1778714762508}],"ts":1778714762508},{"id":"yzq48lfdj","date":"05/05/2026","client":"PEPE","type":"compra","usdtAmount":32480.63,"tcClient":1475.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"kkxs79ed9","amount":2254432.52,"ts":1778714762508},{"id":"ofka54sib","amount":1400000.0,"ts":1778714762508},{"id":"cvac89e1h","amount":1410000.0,"ts":1778714762509},{"id":"fz1ts8w74","amount":1500000.0,"ts":1778714762509},{"id":"fp09e7o20","amount":577500.0,"ts":1778714762509},{"id":"4ze91z0il","amount":1415000.0,"ts":1778714762509},{"id":"bolxggzch","amount":1415000.0,"ts":1778714762509},{"id":"rz5hqly75","amount":1415000.0,"ts":1778714762509},{"id":"mkl20sjxu","amount":972500.0,"ts":1778714762509},{"id":"5cj7rxikw","amount":21150000.0,"ts":1778714762509},{"id":"acmx9qkde","amount":325000.0,"ts":1778714762509},{"id":"mnrkkiukx","amount":574500.0,"ts":1778714762509},{"id":"3vhlgkezc","amount":1960000.0,"ts":1778714762509},{"id":"v0padfgbp","amount":1540000.0,"ts":1778714762509},{"id":"mkc5404yp","amount":10000000.0,"ts":1778714762509}],"ts":1778714762509},{"id":"ub7bwzrq3","date":"06/05/2026","client":"PEPE","type":"compra","usdtAmount":49998.5,"tcClient":1461.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"5rgtoe4v1","amount":1000000.0,"ts":1778714762509},{"id":"g3gyryun7","amount":260000.0,"ts":1778714762509},{"id":"61g8pczwy","amount":562000.0,"ts":1778714762509},{"id":"pt52p2ugj","amount":10000000.0,"ts":1778714762509},{"id":"ch755i1wc","amount":2800000.0,"ts":1778714762509},{"id":"u7z8zuy7a","amount":935805.0,"ts":1778714762509},{"id":"91pyjgy7k","amount":1000000.0,"ts":1778714762509},{"id":"cv9brlrba","amount":142168.0,"ts":1778714762509},{"id":"xphgigui3","amount":800000.0,"ts":1778714762509},{"id":"4rnwn6cw3","amount":2000000.0,"ts":1778714762509},{"id":"l29d9lvu2","amount":1445000.0,"ts":1778714762509},{"id":"86lsyr8hd","amount":278000.0,"ts":1778714762509},{"id":"se0f9e5mc","amount":760000.0,"ts":1778714762509},{"id":"2ig1vx13r","amount":1000000.0,"ts":1778714762509},{"id":"xu3wx8gh5","amount":3688380.0,"ts":1778714762509},{"id":"gazpwu2iv","amount":2230000.0,"ts":1778714762509},{"id":"ufne38tcx","amount":2408000.0,"ts":1778714762509},{"id":"be9nauy39","amount":2600000.0,"ts":1778714762509},{"id":"1kfdul2mf","amount":1050000.0,"ts":1778714762509},{"id":"sit257iln","amount":1000000.0,"ts":1778714762509},{"id":"yajit0h71","amount":1390000.0,"ts":1778714762509},{"id":"m0c9s9uvs","amount":2085000.0,"ts":1778714762509},{"id":"t7ve4p1ui","amount":3174000.0,"ts":1778714762509},{"id":"kurh4ze2s","amount":1380000.0,"ts":1778714762509},{"id":"g1gw03b2l","amount":2895000.0,"ts":1778714762509},{"id":"c6103xdxp","amount":2770000.0,"ts":1778714762509},{"id":"cetaoiilw","amount":4230000.0,"ts":1778714762509},{"id":"2q8j8o5xi","amount":1034155.5,"ts":1778714762509},{"id":"8spjcwvs3","amount":1390000.0,"ts":1778714762509},{"id":"fkb52tkkb","amount":3000000.0,"ts":1778714762509},{"id":"fvnasffy0","amount":3000300.0,"ts":1778714762509},{"id":"agor9a2qi","amount":4110000.0,"ts":1778714762509},{"id":"x6dxc6jhj","amount":2500000.0,"ts":1778714762509},{"id":"1o7rkp376","amount":2000000.0,"ts":1778714762509},{"id":"mk9ozf4tk","amount":2130000.0,"ts":1778714762509}],"ts":1778714762509},{"id":"7gweo3y3q","date":"06/05/2026","client":"PABLITO","type":"venta","usdtAmount":1562.0,"tcClient":1476.89,"tcOtc":null,"bank":"Letsbit","exchange":"Binance","tts":[{"id":"fy8tg482t","amount":2306897.5,"ts":1778714762509}],"ts":1778714762509},{"id":"quopc4xqd","date":"07/05/2026","client":"JAVI","type":"compra","usdtAmount":50000.0,"tcClient":1468.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"l5ducvstg","amount":1450000.0,"ts":1778714762509},{"id":"709hyd81v","amount":3360000.0,"ts":1778714762509},{"id":"7lbe1h10t","amount":3412836.0,"ts":1778714762509},{"id":"mp4u2arxz","amount":2201000.0,"ts":1778714762509},{"id":"lkxtj21vv","amount":1500000.0,"ts":1778714762509},{"id":"h7axz3e43","amount":615000.0,"ts":1778714762509},{"id":"pog2rewd0","amount":2770000.0,"ts":1778714762509},{"id":"k2b8w1riw","amount":980000.0,"ts":1778714762509},{"id":"wp6t1hsag","amount":6925000.0,"ts":1778714762509},{"id":"j7sly1hmz","amount":270000.0,"ts":1778714762509},{"id":"68l7mg5ph","amount":2000000.0,"ts":1778714762509},{"id":"gri2qf6er","amount":1668000.0,"ts":1778714762509},{"id":"n7cd066nb","amount":790000.0,"ts":1778714762509},{"id":"fbul226m2","amount":839750.0,"ts":1778714762509},{"id":"kpf1fgwoy","amount":690000.0,"ts":1778714762509},{"id":"4yqe87kxj","amount":800000.0,"ts":1778714762509},{"id":"blb2y6fby","amount":800000.0,"ts":1778714762509},{"id":"o7rk4eafo","amount":491233.6,"ts":1778714762509},{"id":"vbhlqdm0p","amount":1060000.0,"ts":1778714762509},{"id":"i61bfrldn","amount":1350000.0,"ts":1778714762509},{"id":"j8ee5ntxf","amount":697500.0,"ts":1778714762509},{"id":"7dwuhl5qz","amount":30000000.0,"ts":1778714762509},{"id":"a27mkwh7y","amount":8490000.0,"ts":1778714762509},{"id":"ucb6wydee","amount":239680.4,"ts":1778714762509}],"ts":1778714762509},{"id":"mczmurvyo","date":"07/05/2026","client":"PABLITO","type":"venta","usdtAmount":1562.0,"tcClient":1476.89,"tcOtc":null,"bank":"Letsbit","exchange":"Bybit","tts":[{"id":"izi1g3vhb","amount":2306897.5,"ts":1778714762509}],"ts":1778714762509},{"id":"ykqju86lg","date":"07/05/2026","client":"VERO OP1","type":"compra","usdtAmount":83786.76,"tcClient":1462.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"sbtfdsi69","amount":504000.0,"ts":1778714762509},{"id":"l40z5botu","amount":280000.0,"ts":1778714762509},{"id":"83xpficmx","amount":245518.0,"ts":1778714762509},{"id":"jaf1th451","amount":2641000.0,"ts":1778714762509},{"id":"sprx2h3zk","amount":1027000.0,"ts":1778714762509},{"id":"s7uws7x2m","amount":1450000.0,"ts":1778714762509},{"id":"vb9pnsflw","amount":417000.0,"ts":1778714762509},{"id":"dn1r8zd2z","amount":831000.0,"ts":1778714762509},{"id":"cd48mzdc3","amount":903500.0,"ts":1778714762509},{"id":"jmpv5pgxy","amount":5650000.0,"ts":1778714762509},{"id":"eqw73ciby","amount":1000000.0,"ts":1778714762509},{"id":"87dv1wbzs","amount":444819.0,"ts":1778714762509},{"id":"0tefpci8z","amount":1390000.0,"ts":1778714762509},{"id":"hnpruz0wv","amount":820944.0,"ts":1778714762509},{"id":"brgf5rhy6","amount":330000.0,"ts":1778714762509},{"id":"6nwdd64yk","amount":4200000.0,"ts":1778714762509},{"id":"7drbnz8s1","amount":5000000.0,"ts":1778714762509},{"id":"okzxjh5ov","amount":278000.0,"ts":1778714762509},{"id":"dfj5fcjcl","amount":695000.0,"ts":1778714762509},{"id":"8dv83hiiz","amount":736367.0,"ts":1778714762509},{"id":"k8kudnllp","amount":620000.0,"ts":1778714762509},{"id":"1gvtfdryc","amount":200000.0,"ts":1778714762509},{"id":"0kha799iw","amount":230000.0,"ts":1778714762509},{"id":"414qtprxi","amount":345000.0,"ts":1778714762509},{"id":"30bxcsxzk","amount":22404400.0,"ts":1778714762509},{"id":"z6u5ur4td","amount":1380000.0,"ts":1778714762509},{"id":"li85xtsrw","amount":3000000.0,"ts":1778714762509},{"id":"xgpw9spb1","amount":550000.0,"ts":1778714762509},{"id":"6ran1p393","amount":1760000.0,"ts":1778714762509},{"id":"lh31ga7fg","amount":450000.0,"ts":1778714762509},{"id":"4tz4wpjjs","amount":1939000.0,"ts":1778714762509},{"id":"2246gbbbj","amount":2800000.0,"ts":1778714762509},{"id":"ietl8wgt3","amount":1380000.0,"ts":1778714762509},{"id":"i2ixpvo0h","amount":22404400.0,"ts":1778714762509},{"id":"kw47cu1kb","amount":139000.0,"ts":1778714762509},{"id":"ms0hf2krw","amount":400000.0,"ts":1778714762509},{"id":"l10rwxovt","amount":1940000.0,"ts":1778714762509},{"id":"wghvqfpdx","amount":2085000.0,"ts":1778714762509},{"id":"fn4vugd7y","amount":1000000.0,"ts":1778714762509},{"id":"wtm9zv55q","amount":5611800.0,"ts":1778714762509},{"id":"2menq2fvi","amount":1600000.0,"ts":1778714762509},{"id":"c1arebsm5","amount":1736000.0,"ts":1778714762509},{"id":"sk75srb9a","amount":2820000.0,"ts":1778714762509},{"id":"yxu5ttw0c","amount":1500000.0,"ts":1778714762509},{"id":"glh6ua1me","amount":1400000.0,"ts":1778714762509},{"id":"g8843nclm","amount":1053000.0,"ts":1778714762509},{"id":"9ntj9n5gb","amount":690000.0,"ts":1778714762509},{"id":"fgq1xm77k","amount":1634500.0,"ts":1778714762509},{"id":"fritktg6a","amount":5000000.0,"ts":1778714762509},{"id":"tlqywesw7","amount":2790000.0,"ts":1778714762509},{"id":"ndix7fxwn","amount":2790000.0,"ts":1778714762509}],"ts":1778714762509},{"id":"wayw54zyb","date":"07/05/2026","client":"VERO OP2","type":"compra","usdtAmount":133019.66,"tcClient":1466.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"le2bj1ru7","amount":32000000.0,"ts":1778714762509},{"id":"x5hgu5hb5","amount":24000000.0,"ts":1778714762509},{"id":"90b896jvf","amount":86500000.0,"ts":1778714762509},{"id":"h2admem6f","amount":2760000.0,"ts":1778714762509},{"id":"pp2y7gdgj","amount":1100000.0,"ts":1778714762509},{"id":"a40u4wt5c","amount":2820000.0,"ts":1778714762509},{"id":"rdzwpdib0","amount":1300000.0,"ts":1778714762509},{"id":"svxxt6wde","amount":2000000.0,"ts":1778714762510},{"id":"nsjnkhqlu","amount":2000000.0,"ts":1778714762510},{"id":"kekb5c91v","amount":2573600.0,"ts":1778714762510},{"id":"bp1p6qqla","amount":321290.16,"ts":1778714762510},{"id":"w5x78uyoc","amount":3472000.0,"ts":1778714762510},{"id":"frxndx29f","amount":1100000.0,"ts":1778714762510},{"id":"20h219si8","amount":1385000.0,"ts":1778714762510},{"id":"z3lxtpdnr","amount":2631500.0,"ts":1778714762510},{"id":"eohhy17xf","amount":2340105.0,"ts":1778714762510},{"id":"j2fpnt80r","amount":2363000.0,"ts":1778714762510},{"id":"szvna4lrz","amount":700000.0,"ts":1778714762510},{"id":"h8xfvmq9b","amount":2726752.0,"ts":1778714762510},{"id":"mldjzg9d6","amount":1370000.0,"ts":1778714762510},{"id":"8ylqhdyqi","amount":700000.0,"ts":1778714762510},{"id":"1rpfs8rcw","amount":576000.0,"ts":1778714762510},{"id":"pwmvjwyzz","amount":73000.0,"ts":1778714762510},{"id":"1e4qq5kcl","amount":140000.0,"ts":1778714762510},{"id":"ghdbcedd3","amount":100000.0,"ts":1778714762510},{"id":"fmk1z5801","amount":422500.0,"ts":1778714762510},{"id":"b1yekfsog","amount":3900000.0,"ts":1778714762510},{"id":"p6ehiwvus","amount":1500000.0,"ts":1778714762510},{"id":"13ahcs28p","amount":1500000.0,"ts":1778714762510},{"id":"x30io2luo","amount":1500000.0,"ts":1778714762510},{"id":"bpj13i2pg","amount":590570.0,"ts":1778714762510},{"id":"a3b1fpept","amount":2484000.0,"ts":1778714762510},{"id":"6rr2qd077","amount":420000.0,"ts":1778714762510},{"id":"jihror75v","amount":908000.0,"ts":1778714762510},{"id":"ec2nhtn16","amount":500000.0,"ts":1778714762510},{"id":"wa9mdjhsg","amount":700000.0,"ts":1778714762510},{"id":"u24bhn1a0","amount":424500.0,"ts":1778714762510},{"id":"29vf3lhuj","amount":1000000.0,"ts":1778714762510},{"id":"g06ikp8eu","amount":1380000.0,"ts":1778714762510},{"id":"fgixuuw29","amount":725000.0,"ts":1778714762510}],"ts":1778714762510},{"id":"5rr1d6sgt","date":"08/05/2026","client":"PEPE","type":"compra","usdtAmount":60000.0,"tcClient":1466.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"8biztk10y","amount":980000.0,"ts":1778714762510},{"id":"kdo9slfnp","amount":1200000.0,"ts":1778714762510},{"id":"5vqjnyijd","amount":1050000.0,"ts":1778714762510},{"id":"bb849osmd","amount":1200000.0,"ts":1778714762510},{"id":"yxf8siqne","amount":969500.0,"ts":1778714762510},{"id":"jiymlw7q3","amount":6950000.0,"ts":1778714762510},{"id":"3svvue2zv","amount":1120000.0,"ts":1778714762510},{"id":"vysmlrgl2","amount":1500000.0,"ts":1778714762510},{"id":"s1uielypl","amount":950300.0,"ts":1778714762510},{"id":"tfbtwf7iq","amount":5560000.0,"ts":1778714762510},{"id":"k7ohtfh8z","amount":3000000.0,"ts":1778714762510},{"id":"l6l8mh3ud","amount":2300000.0,"ts":1778714762510},{"id":"uiead76rv","amount":1000000.0,"ts":1778714762510},{"id":"6ie9g9c4c","amount":1400000.0,"ts":1778714762510},{"id":"altzxvnuf","amount":1390000.0,"ts":1778714762510},{"id":"al463ys2c","amount":1395000.0,"ts":1778714762510},{"id":"mesl6h594","amount":1092637.2,"ts":1778714762510},{"id":"396nrpbzi","amount":1370000.0,"ts":1778714762510},{"id":"bf3tetf4g","amount":3755821.0,"ts":1778714762510},{"id":"ebcu6xgkw","amount":5143000.0,"ts":1778714762510},{"id":"i97kg962t","amount":3500000.0,"ts":1778714762510},{"id":"mypgvyp03","amount":4045000.0,"ts":1778714762510},{"id":"0lhd1rcap","amount":22000000.0,"ts":1778714762510},{"id":"uqndb56yx","amount":1890000.0,"ts":1778714762510},{"id":"e6ipvuy7g","amount":1500000.0,"ts":1778714762510},{"id":"2jhhi8qnc","amount":2810000.0,"ts":1778714762510},{"id":"kub8y9lyv","amount":40000.0,"ts":1778714762510},{"id":"c80vz47hc","amount":4170000.0,"ts":1778714762510},{"id":"hgi2yidlf","amount":840000.0,"ts":1778714762510},{"id":"642a6hztx","amount":3838741.8,"ts":1778714762510}],"ts":1778714762510},{"id":"2z14886tj","date":"08/05/2026","client":"VERO","type":"compra","usdtAmount":34054.21,"tcClient":1468.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"op6if14fe","amount":24750129.0,"ts":1778714762510},{"id":"151jpkwfi","amount":1400000.0,"ts":1778714762510},{"id":"itmv32yvk","amount":2000000.0,"ts":1778714762510},{"id":"o6odyj4ji","amount":2787200.0,"ts":1778714762510},{"id":"7za2hgktf","amount":414000.0,"ts":1778714762510},{"id":"9rcxgktvn","amount":2800000.0,"ts":1778714762510},{"id":"dctbcypf0","amount":789800.0,"ts":1778714762510},{"id":"q0t81xrju","amount":380000.0,"ts":1778714762510},{"id":"pdpwtdk8i","amount":396452.11,"ts":1778714762510},{"id":"epdjman7x","amount":1410000.0,"ts":1778714762510},{"id":"awwlua46o","amount":600000.0,"ts":1778714762510},{"id":"kql6lj37n","amount":278000.0,"ts":1778714762510},{"id":"xgslgeefa","amount":2070000.0,"ts":1778714762510},{"id":"086noabq4","amount":495000.0,"ts":1778714762510},{"id":"z1pgouud7","amount":500000.0,"ts":1778714762510},{"id":"9u1lj4vlz","amount":8618000.0,"ts":1778714762510},{"id":"cekecgz7r","amount":45000.0,"ts":1778714762510},{"id":"zrv1np1cb","amount":138000.0,"ts":1778714762510},{"id":"t65hop7yn","amount":120000.0,"ts":1778714762510}],"ts":1778714762510},{"id":"izn07kx2l","date":"08/05/2026","client":"JUAN","type":"compra","usdtAmount":40380.0,"tcClient":1470.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"ky1meu0ni","amount":2800000.0,"ts":1778714762510},{"id":"ga7zvn3z9","amount":2589520.52,"ts":1778714762510},{"id":"n4dfukci8","amount":1688162.0,"ts":1778714762510},{"id":"odrzzg20r","amount":3792537.0,"ts":1778714762510},{"id":"ig27lrzdm","amount":12000000.0,"ts":1778714762510},{"id":"0hxyxr850","amount":900000.0,"ts":1778714762510},{"id":"zrnpyq22r","amount":1400000.0,"ts":1778714762510},{"id":"o7qi9jta1","amount":3000000.0,"ts":1778714762510},{"id":"qhas8n0cc","amount":1000000.0,"ts":1778714762510},{"id":"kv3a6xrnv","amount":270000.0,"ts":1778714762510},{"id":"p0eq4nmjm","amount":1390000.0,"ts":1778714762510},{"id":"za0ytckxq","amount":700000.0,"ts":1778714762510},{"id":"8sadr8oez","amount":9584200.0,"ts":1778714762510},{"id":"gws8k08v8","amount":2800000.0,"ts":1778714762510},{"id":"s0z0nxy6i","amount":2200000.0,"ts":1778714762510},{"id":"5xtvf8j8s","amount":616000.0,"ts":1778714762510},{"id":"eoaftjj72","amount":195000.0,"ts":1778714762510},{"id":"5pec6o2lb","amount":880000.0,"ts":1778714762510},{"id":"duu54kcm5","amount":121140.0,"ts":1778714762510},{"id":"148zl1ql8","amount":11432040.48,"ts":1778714762510}],"ts":1778714762510},{"id":"ws89q1nig","date":"11/05/2026","client":"JAVI","type":"compra","usdtAmount":47418.0,"tcClient":1463.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"zezn9k5ok","amount":39372534.0,"ts":1778714762510},{"id":"2caxoesnj","amount":30000000.0,"ts":1778714762510}],"ts":1778714762510},{"id":"pun7eq5dl","date":"11/05/2026","client":"VERO/JAVI","type":"compra","usdtAmount":34199.73,"tcClient":1462.0,"tcOtc":null,"bank":"Ripio","exchange":"Fiwind","tts":[{"id":"j6jcijcjj","amount":1133084.24,"ts":1778714762510},{"id":"tx4s5uvk1","amount":847000.0,"ts":1778714762510},{"id":"sw115lqa2","amount":95000.0,"ts":1778714762510},{"id":"fmvxcezi1","amount":1420000.0,"ts":1778714762510},{"id":"xtbaarbck","amount":697500.0,"ts":1778714762510},{"id":"lozklpc8f","amount":469500.0,"ts":1778714762510},{"id":"7ac5irb4n","amount":1104000.0,"ts":1778714762510},{"id":"p867epogh","amount":688050.0,"ts":1778714762510},{"id":"tqo1gw517","amount":417000.0,"ts":1778714762510},{"id":"1yb214kdp","amount":2800000.0,"ts":1778714762510},{"id":"ng8lyeg9r","amount":9000000.0,"ts":1778714762510},{"id":"hr4bg8q6c","amount":1385000.0,"ts":1778714762510},{"id":"n6cbvckg4","amount":1000000.0,"ts":1778714762510},{"id":"c4nf24mw8","amount":2487500.0,"ts":1778714762510},{"id":"9kiw0tc5c","amount":4185000.0,"ts":1778714762510},{"id":"34pjijamz","amount":2085000.0,"ts":1778714762510},{"id":"prczdd6ue","amount":690000.0,"ts":1778714762510},{"id":"8xyuuoxnk","amount":1390000.0,"ts":1778714762510},{"id":"gkwp32g75","amount":5143000.0,"ts":1778714762510},{"id":"5yyzd89nx","amount":139000.0,"ts":1778714762510},{"id":"yui3yzzts","amount":828000.0,"ts":1778714762510},{"id":"5mly4ocmf","amount":1860000.0,"ts":1778714762510},{"id":"ckg9bv0jd","amount":2760000.0,"ts":1778714762510},{"id":"gcdhipw76","amount":1405000.0,"ts":1778714762510},{"id":"b7id83k82","amount":690000.0,"ts":1778714762510},{"id":"s32blguby","amount":417000.0,"ts":1778714762510},{"id":"3f97j9zsn","amount":4570500.0,"ts":1778714762510},{"id":"ykmyjmvbd","amount":302995.0,"ts":1778714762510}],"ts":1778714762510},{"id":"vhk951337","date":"11/05/2026","client":"PEPE","type":"compra","usdtAmount":53500.0,"tcClient":1468.0,"tcOtc":null,"bank":"Ripio","exchange":"Bybit","tts":[{"id":"jgycppik3","amount":7000000.0,"ts":1778714762511},{"id":"xb1zt6kxk","amount":2520000.0,"ts":1778714762511},{"id":"s919e6vfp","amount":7700000.0,"ts":1778714762511},{"id":"bztzpmsql","amount":5000000.0,"ts":1778714762511},{"id":"924hj66nf","amount":800000.0,"ts":1778714762511},{"id":"4hsvwhf6m","amount":500000.0,"ts":1778714762511},{"id":"8yqpugwot","amount":530200.0,"ts":1778714762511},{"id":"lvwec7jw4","amount":600000.0,"ts":1778714762511},{"id":"4cpnk3tbn","amount":5100000.0,"ts":1778714762511},{"id":"gtb6mxvfd","amount":576400.0,"ts":1778714762511},{"id":"5fyxdbayc","amount":552000.0,"ts":1778714762511},{"id":"ag015tc0s","amount":445300.0,"ts":1778714762511},{"id":"ycqgpsmwm","amount":333800.0,"ts":1778714762511},{"id":"r4srs6nyn","amount":975600.0,"ts":1778714762511},{"id":"i6d5fkaz1","amount":5520000.0,"ts":1778714762511},{"id":"d66toxyxu","amount":600000.0,"ts":1778714762511},{"id":"6k4d5f5s6","amount":498000.0,"ts":1778714762511},{"id":"kdk2i11am","amount":570000.0,"ts":1778714762511},{"id":"4dk2c0pmc","amount":2000000.0,"ts":1778714762511},{"id":"ug9dx72md","amount":5640000.0,"ts":1778714762511},{"id":"zstf709bb","amount":1000000.0,"ts":1778714762511},{"id":"w5r4e1iu3","amount":4845000.0,"ts":1778714762511},{"id":"jrn50jdhn","amount":25000000.0,"ts":1778714762511},{"id":"ms3qp42b5","amount":228423.65,"ts":1778714762511},{"id":"6u93zit98","amount":3276.35,"ts":1778714762511}],"ts":1778714762511},{"id":"msvzx0unr","date":"12/05/2026","client":"JAVI","type":"compra","usdtAmount":27626.2,"tcClient":1466.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"40fvfjfkb","amount":1385000.0,"ts":1778714762511},{"id":"ecm78fre6","amount":1600000.0,"ts":1778714762511},{"id":"pdcgny31x","amount":120000.0,"ts":1778714762511},{"id":"98f9okbs8","amount":562000.0,"ts":1778714762511},{"id":"uqa45f9ze","amount":2200000.0,"ts":1778714762511},{"id":"sb5ypmyg8","amount":469337.0,"ts":1778714762511},{"id":"s18fku31x","amount":1828000.0,"ts":1778714762511},{"id":"st2dp87z6","amount":4794000.0,"ts":1778714762511},{"id":"2qgf19rh3","amount":27541800.0,"ts":1778714762511}],"ts":1778714762511},{"id":"vk0czyw03","date":"13/05/2026","client":"LORE","type":"compra","usdtAmount":46332.72,"tcClient":1467.0,"tcOtc":null,"bank":"Ripio","exchange":"Cocos","tts":[{"id":"y4lb7f4mk","amount":16000000.0,"ts":1778714762511},{"id":"kxumvf916","amount":15300000.0,"ts":1778714762511},{"id":"rcaad1wma","amount":9500000.0,"ts":1778714762511},{"id":"2liloshs0","amount":13000000.0,"ts":1778714762511},{"id":"3cy6sfd54","amount":12500000.0,"ts":1778714762511},{"id":"vulr5u4z4","amount":1669997.0,"ts":1778714762511}],"ts":1778714762511}];

// ═══════════════════════════════════════════════════════════════════
//  STORAGE
// ═══════════════════════════════════════════════════════════════════
const KEY_OPS = "plexo-ops-v1";
const KEY_PREFS = "plexo-prefs-v1";

// Storage: localStorage + Google Drive auto-sync
const loadOps = async () => {
  try {
    const stored = localStorage.getItem(KEY_OPS);
    const arr = stored ? JSON.parse(stored) : [];
    const storedMap = new Map(arr.map(o => [o.id, o]));
    const initialIds = new Set(INITIAL_DATA.map(o => o.id));
    const merged = INITIAL_DATA.map(o => storedMap.has(o.id) ? storedMap.get(o.id) : o);
    const newOps = arr.filter(o => !initialIds.has(o.id));
    return [...merged, ...newOps];
  } catch { return [...INITIAL_DATA]; }
};
const saveOps = async data => {
  try { localStorage.setItem(KEY_OPS, JSON.stringify(data)); } catch {}
  // Drive sync (async, no bloquea UI)
  driveSync.queueSync(data);
};
const loadPrefs = async () => {
  try { return JSON.parse(localStorage.getItem(KEY_PREFS) || "{}"); }
  catch { return {}; }
};
const savePrefs = async p => { try { localStorage.setItem(KEY_PREFS, JSON.stringify(p)); } catch {} };

// ─── GOOGLE DRIVE INTEGRATION ───
const driveSync = (() => {
  const FOLDER_NAME = "Plexo_Backup";
  const FILE_NAME = "plexo_data.json";
  const XLSX_FILE_NAME = "plexo_operaciones.xlsx";
  const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || "";
  const SCOPE = "https://www.googleapis.com/auth/drive.file";

  let tokenClient = null;
  let accessToken = localStorage.getItem("plexo-drive-token") || null;
  let folderId = localStorage.getItem("plexo-drive-folder") || null;
  let fileId = localStorage.getItem("plexo-drive-fileid") || null;
  let xlsxFileId = localStorage.getItem("plexo-drive-xlsxid") || null;
  let queueTimer = null;
  let pendingData = null;
  let listeners = [];

  const notify = (state) => listeners.forEach(l => l(state));
  const onStatus = (fn) => { listeners.push(fn); return () => { listeners = listeners.filter(l => l !== fn); }; };

  const init = () => new Promise((resolve) => {
    if (!CLIENT_ID) { console.warn("CLIENT_ID no configurado"); resolve(false); return; }
    if (window.google?.accounts?.oauth2) {
      tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID, scope: SCOPE,
        callback: (resp) => {
          if (resp.access_token) {
            accessToken = resp.access_token;
            localStorage.setItem("plexo-drive-token", accessToken);
            notify({ connected: true });
            if (pendingData) { syncNow(pendingData); pendingData = null; }
          }
        },
      });
      resolve(true);
    } else {
      // Cargar la librería de Google
      const s = document.createElement("script");
      s.src = "https://accounts.google.com/gsi/client";
      s.async = true; s.defer = true;
      s.onload = () => init().then(resolve);
      document.head.appendChild(s);
    }
  });

  const connect = async () => {
    const ok = await init();
    if (!ok || !tokenClient) { alert("Drive no configurado. Avisame y lo configuramos."); return; }
    tokenClient.requestAccessToken();
  };

  const disconnect = () => {
    accessToken = null; folderId = null; fileId = null; xlsxFileId = null;
    localStorage.removeItem("plexo-drive-token");
    localStorage.removeItem("plexo-drive-folder");
    localStorage.removeItem("plexo-drive-fileid");
    localStorage.removeItem("plexo-drive-xlsxid");
    notify({ connected: false });
  };

  const isConnected = () => !!accessToken;

  const apiCall = async (url, options = {}) => {
    const res = await fetch(url, {
      ...options,
      headers: { ...options.headers, Authorization: `Bearer ${accessToken}` },
    });
    if (res.status === 401) { disconnect(); throw new Error("Sesión expiró, reconectá"); }
    return res;
  };

  const ensureFolder = async () => {
    if (folderId) return folderId;
    // Buscar
    const q = encodeURIComponent(`name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
    const r = await apiCall(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`);
    const d = await r.json();
    if (d.files?.length) { folderId = d.files[0].id; }
    else {
      const c = await apiCall("https://www.googleapis.com/drive/v3/files", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: FOLDER_NAME, mimeType: "application/vnd.google-apps.folder" }),
      });
      const cd = await c.json();
      folderId = cd.id;
    }
    localStorage.setItem("plexo-drive-folder", folderId);
    return folderId;
  };

  const findFile = async (name) => {
    const q = encodeURIComponent(`name='${name}' and '${folderId}' in parents and trashed=false`);
    const r = await apiCall(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`);
    const d = await r.json();
    return d.files?.[0]?.id || null;
  };

  const uploadJson = async (data) => {
    await ensureFolder();
    if (!fileId) fileId = await findFile(FILE_NAME);
    const json = JSON.stringify({ version: 1, ops: data, savedAt: new Date().toISOString() }, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    if (fileId) {
      await apiCall(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`, {
        method: "PATCH", headers: { "Content-Type": "application/json" }, body: blob,
      });
    } else {
      const meta = { name: FILE_NAME, parents: [folderId] };
      const form = new FormData();
      form.append("metadata", new Blob([JSON.stringify(meta)], { type: "application/json" }));
      form.append("file", blob);
      const r = await apiCall("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
        method: "POST", body: form,
      });
      const rd = await r.json();
      fileId = rd.id;
      localStorage.setItem("plexo-drive-fileid", fileId);
    }
  };

  const uploadXlsx = async (data) => {
    await ensureFolder();
    if (!xlsxFileId) xlsxFileId = await findFile(XLSX_FILE_NAME);
    // Generar Excel con TODAS las ops (no por mes)
    const compraRows = [], ventaRows = [];
    const sortByDate = (a, b) => dateToTime(a.date) - dateToTime(b.date);
    [...data].sort(sortByDate).forEach(op => {
      op.tts.forEach(tt => {
        const row = [op.date, "USDT", +(tt.amount / op.tcClient).toFixed(2), op.tcClient, tt.amount, op.exchange || "", op.bank];
        if (op.type === "compra") compraRows.push(row); else ventaRows.push(row);
      });
    });
    const aoa = [
      ["COMPRAS","","","","","","","VENTAS","","","","","",""],
      ["Fecha","Cripto","Cantidad USDT","Tipo de cambio","Cantidad $","Exchange","Banco","Fecha","Cripto","Cantidad USDT","Tipo de cambio","Cantidad $","Exchange","Banco"],
    ];
    const maxLen = Math.max(compraRows.length, ventaRows.length);
    for (let i = 0; i < maxLen; i++) {
      aoa.push([...(compraRows[i] || ["","","","","","",""]), ...(ventaRows[i] || ["","","","","","",""])]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, { s: { r: 0, c: 7 }, e: { r: 0, c: 13 } }];
    ws["!cols"] = Array(14).fill({ wch: 14 });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Operaciones");
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    if (xlsxFileId) {
      await apiCall(`https://www.googleapis.com/upload/drive/v3/files/${xlsxFileId}?uploadType=media`, {
        method: "PATCH", body: blob,
      });
    } else {
      const meta = { name: XLSX_FILE_NAME, parents: [folderId] };
      const form = new FormData();
      form.append("metadata", new Blob([JSON.stringify(meta)], { type: "application/json" }));
      form.append("file", blob);
      const r = await apiCall("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
        method: "POST", body: form,
      });
      const rd = await r.json();
      xlsxFileId = rd.id;
      localStorage.setItem("plexo-drive-xlsxid", xlsxFileId);
    }
  };

  const syncNow = async (data) => {
    if (!isConnected()) { pendingData = data; return; }
    notify({ syncing: true });
    try {
      await uploadJson(data);
      await uploadXlsx(data);
      localStorage.setItem("plexo-drive-lastsync", new Date().toISOString());
      notify({ syncing: false, success: true, lastSync: new Date().toISOString() });
    } catch (e) {
      console.error("Drive sync error:", e);
      notify({ syncing: false, error: e.message });
    }
  };

  // Debounce: agrupa cambios en 3 segundos
  const queueSync = (data) => {
    if (!isConnected()) return;
    clearTimeout(queueTimer);
    queueTimer = setTimeout(() => syncNow(data), 3000);
  };

  // Restaurar desde Drive
  const restoreFromDrive = async () => {
    if (!isConnected()) throw new Error("No conectado a Drive");
    await ensureFolder();
    if (!fileId) fileId = await findFile(FILE_NAME);
    if (!fileId) throw new Error("No hay backup en Drive todavía");
    const r = await apiCall(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);
    const d = await r.json();
    return d.ops || [];
  };

  return { init, connect, disconnect, isConnected, syncNow, queueSync, restoreFromDrive, onStatus, getLastSync: () => localStorage.getItem("plexo-drive-lastsync") };
})();

// ═══════════════════════════════════════════════════════════════════
//  WHATSAPP PARSER (con detección de ambigüedades)
// ═══════════════════════════════════════════════════════════════════
// Parser local con doble pasada (instantáneo, sin tokens)
// Pasada 1: montos con $ obligatorio
// Pasada 2: montos después de "MONTO:" o "IMPORTE:" (sin $)
const parseLocal = (msg) => {
  const amounts = [];
  const warnings = [];
  const collected = [];
  const seenRanges = [];

  // PASADA 1: $ obligatorio
  const re1 = /\$\s*(\d{4,9}(?:,\d{1,2})?|\d{1,3}(?:\.\d{3}){0,2}(?:,\d{1,2})?)(?:\.?-)?/g;
  let m;
  while ((m = re1.exec(msg)) !== null) {
    const raw = m[1];
    if (raw.includes(".")) {
      const parts = raw.split(",")[0].split(".");
      if (!parts.slice(1).every(p => p.length === 3)) continue;
    }
    const value = parseFloat(raw.replace(/\./g, "").replace(",", "."));
    if (value < 1000 || value > 999999999) continue;
    const numStart = m.index + m[0].indexOf(raw);
    collected.push({ value: Math.round(value), idx: numStart });
    seenRanges.push({ start: m.index, end: m.index + m[0].length });
  }

  // PASADA 2: keywords MONTO/IMPORTE/VALOR sin requerir $
  const re2 = /\b(monto|importe|valor)\s*[:=]?\s*\$?\s*(\d{4,9}(?:,\d{1,2})?|\d{1,3}(?:\.\d{3}){0,2}(?:,\d{1,2})?)(?:\.?-)?/gi;
  while ((m = re2.exec(msg)) !== null) {
    const raw = m[2];
    const numStart = m.index + m[0].indexOf(raw);
    // Saltear si ya fue capturado en pasada 1
    if (seenRanges.some(r => numStart >= r.start && numStart < r.end)) continue;
    if (raw.includes(".")) {
      const parts = raw.split(",")[0].split(".");
      if (!parts.slice(1).every(p => p.length === 3)) continue;
    }
    const value = parseFloat(raw.replace(/\./g, "").replace(",", "."));
    if (value < 1000 || value > 999999999) continue;
    collected.push({ value: Math.round(value), idx: numStart });
  }

  // Ordenar por posición en el texto
  collected.sort((a, b) => a.idx - b.idx);

  // Warnings: CBU/CUIT/Alias cerca del monto
  const suspKeywords = /\b(cbu|cuit|cuil|alias|n[°º]\s*de\s*cuenta|cuenta:|titular)\b/i;
  collected.forEach(({ value, idx }) => {
    const before = msg.slice(Math.max(0, idx - 100), idx);
    const lastNewline = before.lastIndexOf("\n");
    const lineBefore = lastNewline >= 0 ? before.slice(lastNewline) : before;
    if (suspKeywords.test(lineBefore)) {
      warnings.push(`⚠ Cerca de $${value.toLocaleString("es-AR")} hay CBU/CUIT/alias — verificá el monto`);
    }
    amounts.push(value);
  });

  if (/\b(usd|usdt|d[oó]lar)\b/i.test(msg) && amounts.length > 0) {
    warnings.push("Detecté mención de USD/USDT — verificá que sea ARS");
  }

  const uniqueWarnings = [...new Set(warnings)].slice(0, 3);
  return { amounts, warnings: uniqueWarnings };
};

// Parser IA como fallback (solo si local no encuentra nada y hay tokens disponibles)
const parseAI = async (msg) => {
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 400,
        messages: [{
          role: "user",
          content: `Extraé montos en pesos argentinos de este mensaje. Respondé SOLO con JSON: {"amounts":[enteros],"warnings":[strings]}. Sin texto extra.\n\n${msg}`,
        }],
      }),
    });
    if (!res.ok) return null;
    const d = await res.json();
    const txt = d?.content?.find(b => b.type === "text")?.text || "{}";
    const clean = txt.replace(/\`\`\`[a-z]*/g, "").replace(/\`\`\`/g, "").trim();
    const parsed = JSON.parse(clean);
    return {
      amounts: Array.isArray(parsed.amounts) ? parsed.amounts.filter(n => typeof n === "number" && n > 0) : [],
      warnings: Array.isArray(parsed.warnings) ? parsed.warnings : [],
    };
  } catch { return null; }
};

const parseWhatsApp = async (msg) => {
  // 1) Intentar parser local primero
  const local = parseLocal(msg);
  if (local.amounts.length > 0) return local;

  // 2) Si local no encuentra nada, intentar IA como fallback
  const ai = await parseAI(msg);
  if (ai && ai.amounts.length > 0) return ai;

  // 3) Nada funcionó
  return { amounts: [], warnings: ["No se detectaron montos válidos en el mensaje"] };
};

// ═══════════════════════════════════════════════════════════════════
//  EXCEL · FORMATO ALASIA
// ═══════════════════════════════════════════════════════════════════
const exportAlasiaXLSX = (ops, monthKey) => {
  const monthOps = ops.filter(o => dateMonthKey(o.date) === monthKey);
  if (!monthOps.length) return false;

  const compraRows = [], ventaRows = [];
  const sortByDate = (a, b) => dateToTime(a.date) - dateToTime(b.date);
  [...monthOps].sort(sortByDate).forEach(op => {
    op.tts.forEach(tt => {
      const usdtForTT = tt.amount / op.tcClient;
      const row = [op.date, "USDT", +usdtForTT.toFixed(2), op.tcClient, tt.amount, op.exchange || "", op.bank];
      if (op.type === "compra") compraRows.push(row);
      else ventaRows.push(row);
    });
  });

  // Header rows
  const aoa = [
    ["COMPRAS", "", "", "", "", "", "", "VENTAS", "", "", "", "", "", ""],
    ["Fecha", "Cripto", "Cantidad USDT", "Tipo de cambio", "Cantidad $", "Exchange", "Banco",
     "Fecha", "Cripto", "Cantidad USDT", "Tipo de cambio", "Cantidad $", "Exchange", "Banco"],
  ];
  const maxLen = Math.max(compraRows.length, ventaRows.length);
  for (let i = 0; i < maxLen; i++) {
    const c = compraRows[i] || ["", "", "", "", "", "", ""];
    const v = ventaRows[i] || ["", "", "", "", "", "", ""];
    aoa.push([...c, ...v]);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
    { s: { r: 0, c: 7 }, e: { r: 0, c: 13 } },
  ];
  ws["!cols"] = Array(14).fill({ wch: 14 });

  // Format date cells as text DD/MM/YYYY
  for (let r = 2; r < aoa.length; r++) {
    [0, 7].forEach(col => {
      const cell = XLSX.utils.encode_cell({ r, c: col });
      if (ws[cell] && ws[cell].v) ws[cell].t = "s";
    });
    // Numbers formatting
    [2, 3, 4, 9, 10, 11].forEach(col => {
      const cell = XLSX.utils.encode_cell({ r, c: col });
      if (ws[cell] && typeof ws[cell].v === "number") {
        ws[cell].t = "n";
        ws[cell].z = "#,##0.00";
      }
    });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Operaciones");
  const [mm, yyyy] = monthKey.split("/");
  XLSX.writeFile(wb, `Plexo_${yyyy}-${mm}.xlsx`);
  return true;
};

// Excel para enviar a un cliente con su op
const exportClientXLSX = (op) => {
  const total = op.usdtAmount * op.tcClient;
  const sent = op.tts.reduce((s, t) => s + t.amount, 0);
  const remaining = total - sent;
  const rows = [
    [`Operación · ${op.client}`],
    [],
    ["Fecha", op.date],
    ["Tipo", op.type === "compra" ? "Compra USDT" : "Venta USDT"],
    ["USDT", op.usdtAmount],
    ["Tipo de cambio", op.tcClient],
    ["Total ARS", total],
    [],
    ["Transferencias enviadas"],
    ["#", "Monto ARS"],
    ...op.tts.map((tt, i) => [i + 1, tt.amount]),
    [],
    ["Total enviado", sent],
    [remaining > 0 ? "Pendiente" : remaining < 0 ? "Sobrante" : "Saldado", Math.abs(remaining)],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 22 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Operación");
  XLSX.writeFile(wb, `Op_${op.client.replace(/\s/g, "_")}_${op.date.replace(/\//g, "-")}.xlsx`);
};

// Backup completo (JSON)
const exportBackup = (ops, prefs) => {
  const data = { version: 1, exportedAt: new Date().toISOString(), ops, prefs };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Plexo_backup_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// ═══════════════════════════════════════════════════════════════════
//  ATOM LOGO
// ═══════════════════════════════════════════════════════════════════
function AtomLogo({ size = 32, animated = false }) {
  const id = `atom-grad-${size}`;
  return (
    <svg width={size} height={size} viewBox="0 0 80 80" style={{ display: "block" }}>
      <defs>
        <linearGradient id={id} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor={C.accent} />
          <stop offset="100%" stopColor={C.accent2} />
        </linearGradient>
      </defs>
      {[0, 45, 90, 135].map(rot => (
        <ellipse key={rot} cx="40" cy="40" rx="32" ry="10" fill="none"
          stroke={`url(#${id})`} strokeWidth="1.8" transform={`rotate(${rot} 40 40)`}
          style={animated ? { animation: `atomspin 8s linear infinite`, transformOrigin: "40px 40px" } : {}} />
      ))}
      <circle cx="40" cy="40" r="5.5" fill={`url(#${id})`} />
    </svg>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  TOAST SYSTEM
// ═══════════════════════════════════════════════════════════════════
function Toast({ toast, onDismiss }) {
  useEffect(() => { const t = setTimeout(onDismiss, toast.undo ? 6000 : 3500); return () => clearTimeout(t); }, [toast.id]); // eslint-disable-line
  const color = toast.kind === "error" ? C.negative : toast.kind === "warn" ? C.warn : C.positive;
  const Icon = toast.kind === "error" ? AlertCircle : toast.kind === "warn" ? AlertCircle : CheckCircle2;
  return (
    <div style={{
      position: "fixed", top: 16, left: "50%", transform: "translateX(-50%)",
      maxWidth: 440, width: "calc(100% - 24px)", zIndex: 200,
      background: "rgba(16,10,28,0.96)", border: `1px solid ${color}40`,
      borderRadius: 12, padding: "11px 14px", display: "flex", alignItems: "center", gap: 10,
      boxShadow: `0 12px 32px rgba(0,0,0,0.5), 0 0 0 1px ${color}20`,
      backdropFilter: "blur(20px)", animation: `slideDown 0.4s ${EASE}`,
    }}>
      <Icon size={16} color={color} />
      <span style={{ flex: 1, color: C.text, fontSize: 13, fontWeight: 500 }}>{toast.message}</span>
      {toast.undo && (
        <button onClick={() => { toast.undo(); onDismiss(); }}
          style={{ background: "transparent", border: `1px solid ${C.accent}40`, color: C.accent, fontFamily: "inherit",
            fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase",
            padding: "5px 10px", borderRadius: 8, cursor: "pointer" }}>
          Deshacer
        </button>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  MAIN APP
// ═══════════════════════════════════════════════════════════════════
export default function App() {
  const [ops, setOps] = useState([]);
  const [prefs, setPrefs] = useState({});
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState("inicio");
  const [sheet, setSheet] = useState(null);
  const [toast, setToast] = useState(null);
  const [showCalc, setShowCalc] = useState(false);

  // Undo/Redo stacks
  const undoStack = useRef([]);
  const redoStack = useRef([]);

  useEffect(() => {
    Promise.all([loadOps(), loadPrefs()]).then(([o, p]) => {
      setOps(o); setPrefs(p); setReady(true);
    });
  }, []);

  useEffect(() => { if (ready) saveOps(ops); }, [ops, ready]);
  useEffect(() => { if (ready) savePrefs(prefs); }, [prefs, ready]);

  // ── Navegación con botón atrás del celular ──
  // Stack de restore: cada acción de navegación pushea una función de "volver"
  const navStack = useRef([]);
  const isPopping = useRef(false);
  const prevSheetRef = useRef(null);
  const prevTabRef = useRef("inicio");

  // Trackear cambios de sheet
  useEffect(() => {
    if (isPopping.current) { prevSheetRef.current = sheet; return; }
    const prev = prevSheetRef.current;
    const curr = sheet;
    if (!prev && curr) {
      // Apertura nueva
      navStack.current.push(() => setSheet(null));
      try { window.history.pushState({ plexo: navStack.current.length }, ""); } catch {}
    } else if (prev && curr && (prev.kind !== curr.kind || prev.id !== curr.id)) {
      // Cambio entre sheets (ej: detail → present)
      const captured = prev;
      navStack.current.push(() => setSheet(captured));
      try { window.history.pushState({ plexo: navStack.current.length }, ""); } catch {}
    } else if (prev && !curr && !isPopping.current) {
      // Cierre manual (botón X) → consumir el history para mantener balance
      try { window.history.back(); } catch {}
    }
    prevSheetRef.current = sheet;
  }, [sheet]);

  // Trackear cambios de tab
  useEffect(() => {
    if (isPopping.current) { prevTabRef.current = tab; return; }
    const prev = prevTabRef.current;
    if (tab !== prev) {
      navStack.current.push(() => setTab(prev));
      try { window.history.pushState({ plexo: navStack.current.length }, ""); } catch {}
    }
    prevTabRef.current = tab;
  }, [tab]);

  // Listener del botón atrás del celular (popstate)
  useEffect(() => {
    const handler = () => {
      if (navStack.current.length > 0) {
        const restore = navStack.current.pop();
        isPopping.current = true;
        try { restore(); } catch {}
        // Reset el flag en el siguiente tick para que los useEffect de tracking no re-pusheen
        setTimeout(() => { isPopping.current = false; }, 50);
      }
    };
    window.addEventListener("popstate", handler);
    // Estado inicial: marcar la "raíz" para que el primer atrás físico tenga algo que consumir
    try {
      if (!window.history.state || window.history.state.plexo === undefined) {
        window.history.replaceState({ plexo: 0 }, "");
      }
    } catch {}
    return () => window.removeEventListener("popstate", handler);
  }, []);

  // ── Toast helpers ──
  const notify = useCallback((message, opts = {}) => {
    setToast({ id: uid(), message, kind: opts.kind || "success", undo: opts.undo });
  }, []);
  const dismissToast = useCallback(() => setToast(null), []);

  // ── State actions with undo support ──
  const pushUndo = useCallback((action) => {
    undoStack.current.push(action);
    if (undoStack.current.length > 20) undoStack.current.shift();
    redoStack.current = [];
  }, []);

  const addOp = useCallback((data) => {
    const newOp = { ...data, id: uid(), tts: [], notes: data.notes || "", ts: Date.now(), createdAt: Date.now() };
    setOps(p => {
      const next = [...p, newOp];
      pushUndo({ kind: "addOp", op: newOp });
      return next;
    });
    setSheet(null);
    notify(`✓ Operación creada con ${data.client}`, { undo: () => {
      setOps(p => p.filter(o => o.id !== newOp.id));
      notify("Operación deshecha");
    }});
  }, [pushUndo, notify]);

  const updateOp = useCallback((id, updates) => {
    setOps(p => {
      const prev = p.find(o => o.id === id);
      if (prev) pushUndo({ kind: "updateOp", id, prev: { ...prev } });
      return p.map(o => o.id === id ? { ...o, ...updates } : o);
    });
  }, [pushUndo]);

  const addTTs = useCallback((opId, amounts) => {
    setOps(p => p.map(o => {
      if (o.id !== opId) return o;
      const newTTs = amounts.map(a => ({ id: uid(), amount: a, ts: Date.now(), verified: false }));
      pushUndo({ kind: "addTTs", opId, addedIds: newTTs.map(t => t.id) });
      return { ...o, tts: [...o.tts, ...newTTs] };
    }));
    const count = amounts.length;
    notify(`✓ ${count} TT${count > 1 ? "s" : ""} agregada${count > 1 ? "s" : ""}`, {
      undo: () => {
        setOps(p => p.map(o => o.id === opId ? { ...o, tts: o.tts.slice(0, o.tts.length - count) } : o));
        notify("TTs deshechas");
      },
    });
  }, [pushUndo, notify]);

  const delTT = useCallback((opId, ttId) => {
    let removed = null;
    setOps(p => p.map(o => {
      if (o.id !== opId) return o;
      const idx = o.tts.findIndex(t => t.id === ttId);
      removed = o.tts[idx];
      return { ...o, tts: o.tts.filter(t => t.id !== ttId) };
    }));
    notify(`TT de ${fmtARS(removed?.amount || 0)} eliminada`, {
      undo: () => {
        setOps(p => p.map(o => o.id === opId ? { ...o, tts: [...o.tts, removed] } : o));
        notify("TT restaurada");
      },
    });
  }, [notify]);

  const delOp = useCallback((id) => {
    const removed = ops.find(o => o.id === id);
    setOps(p => p.filter(o => o.id !== id));
    setSheet(null);
    notify(`Operación de ${removed?.client} eliminada`, {
      undo: () => {
        setOps(p => [...p, removed]);
        notify("Operación restaurada");
      },
    });
  }, [ops, notify]);

  const duplicateOp = useCallback((id) => {
    const orig = ops.find(o => o.id === id);
    if (!orig) return;
    const dup = { ...orig, id: uid(), date: todayStr(), tts: [], ts: Date.now(), createdAt: Date.now() };
    setOps(p => [...p, dup]);
    setSheet({ kind: "detail", id: dup.id });
    notify(`✓ Operación duplicada para ${orig.client}`);
  }, [ops, notify]);

  // ── Memoized data ──
  const today = todayStr();
  const todayOps = useMemo(() => ops.filter(o => o.date === today), [ops, today]);
  const currentMonth = today.slice(3);
  const monthOps = useMemo(() => ops.filter(o => dateMonthKey(o.date) === currentMonth), [ops, currentMonth]);
  const months = useMemo(() => availableMonths(ops), [ops]);

  // Client history for autocomplete
  const clientHistory = useMemo(() => {
    const map = {};
    [...ops].sort((a, b) => b.ts - a.ts).forEach(op => {
      if (!map[op.client]) map[op.client] = { bank: op.bank, exchange: op.exchange, type: op.type };
    });
    return map;
  }, [ops]);

  const detailOp = sheet?.kind === "detail" ? ops.find(o => o.id === sheet.id) : null;
  const presentOp = sheet?.kind === "present" ? ops.find(o => o.id === sheet.id) : null;

  // Total pending today
  const totalPendingToday = useMemo(() => todayOps.reduce((s, op) => {
    const tot = op.usdtAmount * op.tcClient;
    const sent = op.tts.reduce((x, t) => x + t.amount, 0);
    return s + Math.max(0, tot - sent);
  }, 0), [todayOps]);

  if (!ready) return <Loader />;

  return (
    <div style={{
      background: C.bg, minHeight: "100vh",
      fontFamily: "'Outfit', system-ui, sans-serif",
      color: C.text, paddingBottom: 92, maxWidth: 480, margin: "0 auto",
      position: "relative", overflowX: "hidden",
    }}>
      <Styles />
      <BgGrid />

      <div style={{ position: "relative", zIndex: 1 }}>
        {tab === "inicio" && <InicioScreen
          ops={ops} todayOps={todayOps} today={today} totalPending={totalPendingToday}
          onNew={() => setSheet({ kind: "new" })}
          onDetail={id => setSheet({ kind: "detail", id })}
          onPresent={id => setSheet({ kind: "present", id })}
          onQuickAddTT={id => setSheet({ kind: "quickTT", id })}
          onDelete={delOp}
        />}
        {tab === "ganancias" && <GananciasScreen
          ops={ops} months={months} currentMonth={currentMonth}
          onUpdateOp={updateOp}
          onDetail={id => setSheet({ kind: "detail", id })}
          notify={notify}
        />}
        {tab === "buscar" && <BuscarScreen
          ops={ops}
          onDetail={id => setSheet({ kind: "detail", id })}
        />}
      </div>

      <BottomNav tab={tab} onChange={setTab} onSettings={() => setSheet({ kind: "settings" })} />

      <FloatingCalc onClick={() => setShowCalc(true)} />

      {sheet?.kind === "new" && <NewOpSheet
        onClose={() => setSheet(null)} onSave={addOp}
        clientHistory={clientHistory}
      />}
      {detailOp && <OpDetailSheet
        op={detailOp} onClose={() => setSheet(null)}
        onAddTTs={addTTs} onDelTT={delTT} onDelOp={delOp}
        onUpdate={updateOp} onDuplicate={duplicateOp}
        onPresent={() => setSheet({ kind: "present", id: detailOp.id })}
        notify={notify}
      />}
      {presentOp && <PresentSheet op={presentOp} onClose={() => setSheet(null)} notify={notify} />}
      {sheet?.kind === "settings" && <SettingsSheet
        ops={ops} prefs={prefs}
        onClose={() => setSheet(null)}
        onImport={(data) => {
          if (data.ops) setOps(data.ops);
          if (data.prefs) setPrefs(data.prefs);
          notify("✓ Datos restaurados");
        }}
        notify={notify}
      />}
      {sheet?.kind === "quickTT" && (() => {
        const op = ops.find(o => o.id === sheet.id);
        if (!op) return null;
        return <QuickTTSheet op={op} onClose={() => setSheet(null)} onAddTTs={addTTs} notify={notify} />;
      })()}

      {showCalc && <CalcSheet onClose={() => setShowCalc(false)} />}

      {toast && <Toast toast={toast} onDismiss={dismissToast} />}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  STYLES + BG
// ═══════════════════════════════════════════════════════════════════
function Styles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap');
      * { box-sizing: border-box; }
      *::-webkit-scrollbar { display: none; }
      html, body { background: ${C.bg}; }
      input, textarea, button, select { outline: none; font-family: inherit; }
      button { cursor: pointer; border: none; background: none; color: inherit; }
      input::placeholder, textarea::placeholder { color: ${C.muted}; }
      @keyframes spin { to { transform: rotate(360deg); } }
      @keyframes atomspin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      @keyframes fadeUp { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: translateY(0); } }
      @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      @keyframes slideDown { from { opacity: 0; transform: translate(-50%, -16px); } to { opacity: 1; transform: translate(-50%, 0); } }
      @keyframes slideUp { from { transform: translate(-50%, 100%); } to { transform: translate(-50%, 0); } }
      @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }
      @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
    `}</style>
  );
}
function BgGrid() {
  return (
    <>
      <div style={{
        position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none",
        backgroundImage: `linear-gradient(rgba(232,76,193,0.07) 1px,transparent 1px),linear-gradient(90deg,rgba(232,76,193,0.07) 1px,transparent 1px)`,
        backgroundSize: "20px 20px",
      }} />
      <div style={{
        position: "fixed", top: "-20%", left: "50%", transform: "translateX(-50%)",
        width: 700, height: 500, zIndex: 0, pointerEvents: "none",
        background: `radial-gradient(ellipse, rgba(232,76,193,0.05) 0%, transparent 65%)`,
      }} />
    </>
  );
}
function Loader() {
  return (
    <div style={{ background: C.bg, height: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 18 }}>
      <Styles />
      <AtomLogo size={48} animated />
      <div style={{ fontSize: 10, color: C.muted, letterSpacing: "0.3em", textTransform: "uppercase", fontWeight: 600 }}>Plexo</div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SCREEN · INICIO
// ═══════════════════════════════════════════════════════════════════
function InicioScreen({ ops, todayOps, today, totalPending, onNew, onDetail, onPresent, onQuickAddTT, onDelete }) {
  const [showHistory, setShowHistory] = useState(false);
  const [search, setSearch] = useState("");
  const [showSearch, setShowSearch] = useState(false);

  const past = ops.filter(o => o.date !== today);
  const byDate = past.reduce((acc, o) => { (acc[o.date] = acc[o.date] || []).push(o); return acc; }, {});
  const pastDates = Object.keys(byDate).sort((a, b) => dateToTime(b) - dateToTime(a));

  // Quick search filter
  const filterFn = (op) => {
    if (!search.trim()) return true;
    const q = search.toLowerCase().trim();
    return op.client.toLowerCase().includes(q) ||
           op.bank?.toLowerCase().includes(q) ||
           op.exchange?.toLowerCase().includes(q) ||
           op.notes?.toLowerCase().includes(q);
  };
  const filteredToday = todayOps.filter(filterFn);

  const [dd, mm] = today.split("/");

  return (
    <div>
      <header style={{ padding: "48px 18px 16px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
              <AtomLogo size={20} />
              <span style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.24em", fontWeight: 700 }}>
                {dayName(today)} · {dd}/{mm}
              </span>
            </div>
            <h1 style={{ fontSize: 30, fontWeight: 800, letterSpacing: "-0.03em", margin: 0, lineHeight: 1 }}>
              Operaciones
            </h1>
          </div>
          <button onClick={() => setShowSearch(s => !s)} style={{
            width: 38, height: 38, borderRadius: 12,
            background: showSearch ? C.accentDim : "transparent",
            border: `1px solid ${showSearch ? C.accent : C.border}`,
            color: showSearch ? C.accent : C.soft,
            display: "flex", alignItems: "center", justifyContent: "center",
            transition: `all 0.2s ${EASE}`,
          }}>
            <Search size={16} />
          </button>
        </div>

        {showSearch && (
          <div style={{ marginTop: 12, animation: "fadeUp 0.3s ease" }}>
            <input
              autoFocus value={search} onChange={e => setSearch(e.target.value)}
              placeholder="Buscar por cliente, banco, nota…"
              style={{
                width: "100%", padding: "11px 14px", borderRadius: 11,
                background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`,
                color: C.text, fontSize: 14,
              }}
            />
          </div>
        )}

        {totalPending > 0 && !showSearch && (
          <div style={{ marginTop: 12, display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: C.warn }}>
            <span style={{ display: "inline-block", width: 6, height: 6, borderRadius: "50%", background: C.warn, boxShadow: `0 0 8px ${C.warn}`, animation: "pulse 2s infinite" }} />
            <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>
              Pendiente · {fmtARS(totalPending)}
            </span>
          </div>
        )}
      </header>

      <div style={{ padding: "0 16px" }}>
        <NewOpButton onClick={onNew} />

        {filteredToday.length === 0 ? (
          search ? <Empty label="Sin resultados" />
                 : <Empty label="Sin operaciones hoy" cta="Tocá + Nueva operación para empezar" />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 9, marginBottom: 24, animation: "fadeUp 0.3s ease" }}>
            {filteredToday.map(op => (
              <OpCard key={op.id} op={op}
                onClick={() => onDetail(op.id)}
                onQuickAddTT={() => onQuickAddTT(op.id)}
                onPresent={() => onPresent(op.id)}
                onDelete={() => onDelete(op.id)} />
            ))}
          </div>
        )}

        <button onClick={() => setShowHistory(h => !h)} style={{
          display: "flex", alignItems: "center", gap: 10, color: C.soft, fontSize: 11,
          fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.16em",
          marginBottom: 16, width: "100%", padding: "8px 0",
        }}>
          <div style={{ flex: 1, height: 1, background: C.border }} />
          <span>Historial · {pastDates.length} día{pastDates.length !== 1 ? "s" : ""}</span>
          <div style={{ flex: 1, height: 1, background: C.border }} />
          <ChevronDown size={12} style={{ transform: showHistory ? "rotate(180deg)" : "none", transition: `transform 0.3s ${EASE}` }} />
        </button>

        {showHistory && (
          pastDates.length === 0
            ? <Empty label="Sin historial" />
            : pastDates.map(date => (
              <div key={date} style={{ marginBottom: 18, animation: "fadeUp 0.3s ease" }}>
                <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.16em", marginBottom: 8, display: "flex", alignItems: "center", gap: 10 }}>
                  <span>{date.slice(0, 5)}</span>
                  <div style={{ flex: 1, height: 1, background: C.border }} />
                  <span>{byDate[date].length} op{byDate[date].length !== 1 ? "s" : ""}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {byDate[date].filter(filterFn).map(op => (
                    <OpCard key={op.id} op={op}
                      onClick={() => onDetail(op.id)}
                      onQuickAddTT={() => onQuickAddTT(op.id)}
                      onPresent={() => onPresent(op.id)}
                      onDelete={() => onDelete(op.id)} />
                  ))}
                </div>
              </div>
            ))
        )}
      </div>
    </div>
  );
}

function NewOpButton({ onClick }) {
  return (
    <button onClick={onClick} style={{
      width: "100%", padding: "14px", borderRadius: 13,
      background: `linear-gradient(135deg, ${C.accent}, ${C.accent2}cc)`,
      color: C.bg, fontSize: 12, fontWeight: 800,
      letterSpacing: "0.14em", textTransform: "uppercase",
      marginBottom: 16, position: "relative", overflow: "hidden",
      boxShadow: `0 8px 32px rgba(232,76,193,0.3)`,
      transition: `transform 0.2s ${EASE}`,
    }}
    onTouchStart={e => e.currentTarget.style.transform = "scale(0.98)"}
    onTouchEnd={e => e.currentTarget.style.transform = "scale(1)"}>
      <span style={{ position: "relative", zIndex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
        <Plus size={16} strokeWidth={2.5} /> Nueva operación
      </span>
      <div style={{ position: "absolute", inset: 0, background: "linear-gradient(135deg, transparent 40%, rgba(255,255,255,0.1))" }} />
    </button>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  OP CARD (con swipe + quick actions)
// ═══════════════════════════════════════════════════════════════════
function OpCard({ op, onClick, onQuickAddTT, onPresent, onDelete }) {
  const [swipeX, setSwipeX] = useState(0);
  const [touchStart, setTouchStart] = useState(null);
  const [confirmDel, setConfirmDel] = useState(false);
  const tc = T[op.type];

  const total = op.usdtAmount * op.tcClient;
  const sent = op.tts.reduce((s, t) => s + t.amount, 0);
  const remaining = total - sent;
  const pct = total > 0 ? Math.min(100, Math.max(0, (sent / total) * 100)) : 0;
  const done = Math.abs(remaining) < 1;
  const overshoot = remaining < 0 && !done;
  const nearComplete = !done && !overshoot && remaining > 0 && remaining < NEAR_COMPLETE;

  const onTouchStart = e => { setTouchStart(e.touches[0].clientX); setConfirmDel(false); };
  const onTouchMove = e => {
    if (touchStart === null) return;
    const dx = e.touches[0].clientX - touchStart;
    if (dx < 0) setSwipeX(Math.max(dx, -120));
    else if (swipeX < 0) setSwipeX(Math.min(0, swipeX + dx * 0.5));
  };
  const onTouchEnd = () => {
    if (swipeX < -70) setSwipeX(-100);
    else setSwipeX(0);
    setTouchStart(null);
  };

  const handleCardClick = e => {
    if (swipeX !== 0) { setSwipeX(0); return; }
    onClick();
  };

  const handleDeleteClick = (e) => {
    e.stopPropagation();
    if (!confirmDel) { setConfirmDel(true); return; }
    onDelete();
    setSwipeX(0);
    setConfirmDel(false);
  };

  return (
    <div style={{ position: "relative", overflow: "hidden", borderRadius: 14 }}>
      {/* Delete action revealed by swipe */}
      <div style={{
        position: "absolute", right: 0, top: 0, bottom: 0, width: 100,
        background: confirmDel ? C.negative : C.negativeDim,
        display: "flex", alignItems: "center", justifyContent: "center",
        gap: 6, color: confirmDel ? "#fff" : C.negative,
        fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase",
        transition: `background 0.2s ${EASE}`,
        cursor: "pointer",
      }} onClick={handleDeleteClick}>
        <Trash2 size={15} />
        {confirmDel ? "Confirmar" : "Eliminar"}
      </div>

      <div onClick={handleCardClick}
        onTouchStart={onTouchStart} onTouchMove={onTouchMove} onTouchEnd={onTouchEnd}
        style={{
          background: C.card,
          border: `1px solid ${done ? tc.border : nearComplete ? `${C.warn}44` : overshoot ? `${C.warn}66` : C.border}`,
          borderRadius: 14, padding: "12px 14px",
          backdropFilter: "blur(20px)",
          boxShadow: done ? tc.glow : "0 2px 12px rgba(0,0,0,0.3)",
          transform: `translateX(${swipeX}px)`,
          transition: touchStart === null ? `transform 0.3s ${EASE}` : "none",
          position: "relative",
        }}>
        <div style={{ position: "absolute", top: 0, left: 0, width: 3, bottom: 0, background: `linear-gradient(180deg, ${tc.primary}, ${tc.primary}55)`, borderRadius: "3px 0 0 3px" }} />

        <div style={{ marginLeft: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 2 }}>
                <span style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{op.client}</span>
                <span style={{ fontSize: 8, fontWeight: 800, padding: "2px 6px", borderRadius: 4, background: tc.bg, color: tc.primary, letterSpacing: "0.12em" }}>
                  {tc.label}
                </span>
                {nearComplete && <span style={{ fontSize: 8, fontWeight: 800, padding: "2px 6px", borderRadius: 4, background: C.warnDim, color: C.warn, letterSpacing: "0.1em" }}>CASI</span>}
                {overshoot && <span style={{ fontSize: 8, fontWeight: 800, padding: "2px 6px", borderRadius: 4, background: C.warnDim, color: C.warn, letterSpacing: "0.1em" }}>+EXC</span>}
              </div>
              <div style={{ fontSize: 11, color: C.muted }}>
                {op.exchange || "—"} · {op.bank}
                {op.notes && <span style={{ marginLeft: 8, color: C.accent2 }}>· nota</span>}
              </div>
            </div>
            <div style={{ textAlign: "right", marginLeft: 8 }}>
              <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, fontWeight: 700, color: tc.primary }}>
                {fmtUSDT(op.usdtAmount)}
              </div>
              <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>TC {op.tcClient.toLocaleString("es-AR")}</div>
            </div>
          </div>

          <div style={{ height: 3, background: "rgba(255,255,255,0.04)", borderRadius: 2, marginBottom: 7, overflow: "hidden" }}>
            <div style={{
              height: "100%", width: `${pct}%`,
              background: `linear-gradient(90deg, ${tc.primary}66, ${tc.primary})`,
              borderRadius: 2, transition: `width 0.5s ${EASE}`,
            }} />
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 11 }}>
            <span style={{ color: C.muted }}>
              Env. <span style={{ color: C.text, fontFamily: "'JetBrains Mono', monospace" }}>{fmtARS(sent)}</span>
            </span>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{
                color: done ? tc.primary : overshoot ? C.warn : nearComplete ? C.warn : C.accent2,
                fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, fontSize: 11,
              }}>
                {done ? "✓ Completo" : overshoot ? `+${fmtARS(-remaining)}` : `−${fmtARS(remaining)}`}
              </span>
              {!done && (
                <button onClick={e => { e.stopPropagation(); onQuickAddTT(); }} style={{
                  width: 26, height: 26, borderRadius: 8,
                  background: C.accentDim, border: `1px solid ${C.accent}44`, color: C.accent,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  transition: `all 0.2s ${EASE}`,
                }} title="Agregar TT rápido">
                  <Zap size={12} strokeWidth={2.5} />
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SCREEN · GANANCIAS
// ═══════════════════════════════════════════════════════════════════
function GananciasScreen({ ops, months, currentMonth, onUpdateOp, onDetail, notify }) {
  const [selMonth, setSelMonth] = useState(currentMonth);
  const [showStats, setShowStats] = useState(false);
  const [showByClient, setShowByClient] = useState(false);

  const monthOps = useMemo(() => ops.filter(o => dateMonthKey(o.date) === selMonth), [ops, selMonth]);

  const enriched = monthOps.map(op => ({
    ...op,
    profit: op.tcOtc ? (op.type === "compra" ? op.tcClient - op.tcOtc : op.tcOtc - op.tcClient) * op.usdtAmount : null,
  }));
  const totalProfit = enriched.filter(o => o.profit !== null).reduce((s, o) => s + o.profit, 0);
  const totalUSDT = monthOps.reduce((s, o) => s + o.usdtAmount, 0);
  const missing = monthOps.filter(o => !o.tcOtc).length;

  // Stats
  const daysOperated = new Set(monthOps.map(o => o.date)).size;
  const topClient = useMemo(() => {
    const map = {};
    enriched.forEach(o => {
      map[o.client] = map[o.client] || { count: 0, usdt: 0, profit: 0 };
      map[o.client].count++;
      map[o.client].usdt += o.usdtAmount;
      if (o.profit !== null) map[o.client].profit += o.profit;
    });
    return Object.entries(map).sort((a, b) => b[1].usdt - a[1].usdt)[0];
  }, [enriched]);
  const avgUSDT = monthOps.length ? totalUSDT / monthOps.length : 0;

  const byClient = useMemo(() => {
    const map = {};
    enriched.forEach(o => {
      map[o.client] = map[o.client] || { count: 0, usdt: 0, profit: 0, ars: 0 };
      map[o.client].count++;
      map[o.client].usdt += o.usdtAmount;
      map[o.client].ars += o.usdtAmount * o.tcClient;
      if (o.profit !== null) map[o.client].profit += o.profit;
    });
    return Object.entries(map).sort((a, b) => b[1].profit - a[1].profit);
  }, [enriched]);

  const handleExport = () => {
    const ok = exportAlasiaXLSX(ops, selMonth);
    if (ok) notify(`✓ Excel ${selMonth} descargado`);
    else notify("No hay operaciones ese mes", { kind: "warn" });
  };

  return (
    <div>
      <header style={{ padding: "48px 18px 16px" }}>
        <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.24em", marginBottom: 6, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}>
          <TrendingUp size={11} color={C.positive} />
          {monthLabel(`01/${selMonth}`)}
        </div>
        <h1 style={{ fontSize: 30, fontWeight: 800, letterSpacing: "-0.03em", margin: 0, lineHeight: 1 }}>Ganancias</h1>
      </header>

      <div style={{ padding: "0 16px" }}>
        {/* Month selector */}
        {months.length > 1 && (
          <div style={{ display: "flex", gap: 6, overflowX: "auto", marginBottom: 14, paddingBottom: 2 }}>
            {months.map(m => (
              <button key={m} onClick={() => setSelMonth(m)} style={{
                padding: "8px 14px", borderRadius: 10, whiteSpace: "nowrap",
                background: selMonth === m ? C.accentDim : "transparent",
                border: `1px solid ${selMonth === m ? C.accent : C.border}`,
                color: selMonth === m ? C.accent : C.soft,
                fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase",
                fontFamily: "inherit", transition: `all 0.2s ${EASE}`,
              }}>
                {m}
              </button>
            ))}
          </div>
        )}

        {/* Profit hero */}
        <div style={{
          background: C.positiveDim, border: `1px solid ${C.positiveBorder}`,
          borderRadius: 18, padding: 22, marginBottom: 14,
          backdropFilter: "blur(20px)",
          boxShadow: `0 8px 40px rgba(94,234,212,0.06)`,
          position: "relative", overflow: "hidden",
        }}>
          <div style={{ position: "absolute", top: -50, right: -50, width: 160, height: 160, background: `radial-gradient(circle, ${C.positive}22 0%, transparent 70%)`, pointerEvents: "none" }} />
          <div style={{ position: "relative", zIndex: 1 }}>
            <div style={{ fontSize: 10, color: C.positive, textTransform: "uppercase", letterSpacing: "0.22em", marginBottom: 10, fontWeight: 700 }}>
              Ganancia neta del mes
            </div>
            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 30, fontWeight: 700, color: C.positive, lineHeight: 1, letterSpacing: "-0.02em" }}>
              {fmtARS(totalProfit)}
            </div>
            <div style={{ display: "flex", gap: 18, marginTop: 12, fontSize: 11 }}>
              <div>
                <div style={{ color: C.muted, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 2 }}>Volumen</div>
                <div style={{ color: C.text, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{fmtUSDT(totalUSDT)}</div>
              </div>
              <div>
                <div style={{ color: C.muted, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 2 }}>Ops</div>
                <div style={{ color: C.text, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{monthOps.length}</div>
              </div>
              {missing > 0 && (
                <div>
                  <div style={{ color: C.warn, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 2 }}>Sin TC OTC</div>
                  <div style={{ color: C.warn, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{missing}</div>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Stats expandible */}
        <button onClick={() => setShowStats(s => !s)} style={{
          width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "12px 14px", marginBottom: 10, borderRadius: 12,
          background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`,
          color: C.text, fontSize: 12, fontWeight: 600, letterSpacing: "0.08em",
        }}>
          <span style={{ display: "flex", alignItems: "center", gap: 8, textTransform: "uppercase", color: C.soft }}>
            <Sparkles size={13} /> Estadísticas
          </span>
          <ChevronDown size={14} color={C.muted} style={{ transform: showStats ? "rotate(180deg)" : "none", transition: `transform 0.3s ${EASE}` }} />
        </button>
        {showStats && (
          <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 14, marginBottom: 14, backdropFilter: "blur(16px)", animation: "fadeUp 0.3s ease" }}>
            <Stat label="Días operados" value={daysOperated} />
            <Stat label="Promedio USDT por op" value={fmtUSDT(avgUSDT)} />
            <Stat label="Cliente top" value={topClient ? `${topClient[0]} · ${fmtUSDT(topClient[1].usdt)}` : "—"} />
            <Stat label="Ganancia promedio por op" value={monthOps.length ? fmtARS(totalProfit / monthOps.length) : "—"} last />
          </div>
        )}

        {/* By client expandible */}
        <button onClick={() => setShowByClient(s => !s)} style={{
          width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "12px 14px", marginBottom: 10, borderRadius: 12,
          background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`,
          color: C.text, fontSize: 12, fontWeight: 600, letterSpacing: "0.08em",
        }}>
          <span style={{ display: "flex", alignItems: "center", gap: 8, textTransform: "uppercase", color: C.soft }}>
            <Users size={13} /> Por cliente · {byClient.length}
          </span>
          <ChevronDown size={14} color={C.muted} style={{ transform: showByClient ? "rotate(180deg)" : "none", transition: `transform 0.3s ${EASE}` }} />
        </button>
        {showByClient && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14, animation: "fadeUp 0.3s ease" }}>
            {byClient.map(([name, info]) => (
              <div key={name} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{name}</div>
                  <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>{info.count} op · {fmtUSDT(info.usdt)}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, fontWeight: 700, color: info.profit >= 0 ? C.positive : C.negative }}>
                    {fmtARS(info.profit)}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Export */}
        <button onClick={handleExport} style={{
          width: "100%", padding: 13, borderRadius: 12,
          background: C.accentDim, border: `1px solid ${C.accent}40`, color: C.accent,
          fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase",
          marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
        }}>
          <Download size={13} /> Excel Alasia · {selMonth}
        </button>

        {/* Ops list with profit */}
        {monthOps.length === 0 ? <Empty label="Sin operaciones este mes" /> : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {[...enriched].sort((a, b) => dateToTime(b.date) - dateToTime(a.date)).map(op => {
              const tc = T[op.type];
              const spread = op.tcOtc ? ((op.type === "compra" ? op.tcClient - op.tcOtc : op.tcOtc - op.tcClient) / op.tcOtc) * 100 : null;
              return (
                <div key={op.id} onClick={() => onDetail(op.id)} style={{
                  background: C.card, border: `1px solid ${C.border}`, borderRadius: 12,
                  padding: "12px 14px", backdropFilter: "blur(16px)",
                  position: "relative", overflow: "hidden", cursor: "pointer",
                }}>
                  <div style={{ position: "absolute", top: 0, left: 0, width: 3, bottom: 0, background: `linear-gradient(180deg, ${tc.primary}, ${tc.primary}55)` }} />
                  <div style={{ marginLeft: 8, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 2 }}>
                        <span style={{ fontSize: 14, fontWeight: 700 }}>{op.client}</span>
                        <span style={{ fontSize: 8, padding: "2px 6px", borderRadius: 4, background: tc.bg, color: tc.primary, fontWeight: 800, letterSpacing: "0.12em" }}>{tc.label}</span>
                      </div>
                      <div style={{ fontSize: 11, color: C.muted, marginBottom: 3 }}>{op.date.slice(0, 5)} · {op.exchange || "—"}</div>
                      {op.tcOtc
                        ? <div style={{ fontSize: 10, color: C.muted, fontFamily: "'JetBrains Mono', monospace" }}>TC {op.tcClient.toLocaleString("es-AR")} → OTC {op.tcOtc.toLocaleString("es-AR")}</div>
                        : <div style={{ fontSize: 10, color: C.warn }}>⚠ TC OTC pendiente</div>}
                    </div>
                    <div style={{ textAlign: "right" }}>
                      {op.profit !== null
                        ? <>
                            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, fontWeight: 700, color: op.profit >= 0 ? C.positive : C.negative }}>
                              {fmtARS(op.profit)}
                            </div>
                            <div style={{ fontSize: 10, color: C.soft, marginTop: 2 }}>{fmtPct(spread)} · {fmtUSDT(op.usdtAmount)}</div>
                          </>
                        : <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: C.muted }}>—</div>}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
function Stat({ label, value, last }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: last ? "none" : `1px solid ${C.border}` }}>
      <span style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.1em", fontWeight: 600 }}>{label}</span>
      <span style={{ fontSize: 13, color: C.text, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{value}</span>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SCREEN · BUSCAR
// ═══════════════════════════════════════════════════════════════════
function BuscarScreen({ ops, onDetail }) {
  const [query, setQuery] = useState("");
  const [filterType, setFilterType] = useState("all"); // all|compra|venta
  const [filterClient, setFilterClient] = useState("all");
  const [filterBank, setFilterBank] = useState("all");

  const allClients = useMemo(() => [...new Set(ops.map(o => o.client))].sort(), [ops]);
  const allBanks = useMemo(() => [...new Set(ops.map(o => o.bank))].sort(), [ops]);

  const isNumQuery = query.trim() && /^[\d.,\s$]+$/.test(query.trim());
  const queryNum = isNumQuery ? parseNum(query) : null;

  const results = useMemo(() => {
    const out = [];
    for (const op of ops) {
      if (filterType !== "all" && op.type !== filterType) continue;
      if (filterClient !== "all" && op.client !== filterClient) continue;
      if (filterBank !== "all" && op.bank !== filterBank) continue;

      if (!query.trim()) {
        out.push({ kind: "op", op });
        continue;
      }

      if (isNumQuery) {
        // Match TT amount exactly
        op.tts.forEach((tt, idx) => {
          if (Math.round(tt.amount) === Math.round(queryNum)) {
            out.push({ kind: "tt", op, tt, idx });
          }
        });
      } else {
        const q = query.toLowerCase().trim();
        if (op.client.toLowerCase().includes(q) ||
            op.bank?.toLowerCase().includes(q) ||
            op.exchange?.toLowerCase().includes(q) ||
            (op.notes || "").toLowerCase().includes(q)) {
          out.push({ kind: "op", op });
        }
      }
    }
    return out.sort((a, b) => dateToTime(b.op.date) - dateToTime(a.op.date));
  }, [ops, query, filterType, filterClient, filterBank, isNumQuery, queryNum]);

  return (
    <div>
      <header style={{ padding: "48px 18px 16px" }}>
        <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.24em", marginBottom: 6, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}>
          <Search size={11} color={C.accent2} /> Buscar
        </div>
        <h1 style={{ fontSize: 30, fontWeight: 800, letterSpacing: "-0.03em", margin: 0, lineHeight: 1 }}>Buscar</h1>
      </header>

      <div style={{ padding: "0 16px" }}>
        <div style={{ position: "relative", marginBottom: 12 }}>
          <Search size={15} color={C.muted} style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)" }} />
          <input
            autoFocus value={query} onChange={e => setQuery(e.target.value)}
            placeholder="Monto, cliente, nota…"
            style={{
              width: "100%", padding: "13px 14px 13px 38px", borderRadius: 13,
              background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`,
              color: C.text, fontSize: 14,
              fontFamily: isNumQuery ? "'JetBrains Mono', monospace" : "inherit",
            }}
          />
          {query && (
            <button onClick={() => setQuery("")} style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", color: C.muted, padding: 4 }}>
              <X size={15} />
            </button>
          )}
        </div>

        {/* Search hint */}
        {query.trim() && (
          <div style={{ fontSize: 11, color: C.muted, marginBottom: 12, display: "flex", alignItems: "center", gap: 6 }}>
            {isNumQuery
              ? <><Hash size={11} /> Buscando TT por monto exacto: <span style={{ color: C.accent2, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{fmtARS(queryNum)}</span></>
              : <><FileText size={11} /> Buscando texto</>}
          </div>
        )}

        {/* Filters */}
        <div style={{ display: "flex", gap: 6, marginBottom: 14, overflowX: "auto" }}>
          <FilterChip label="Todas" active={filterType === "all"} onClick={() => setFilterType("all")} />
          <FilterChip label="Compras" active={filterType === "compra"} onClick={() => setFilterType("compra")} color={T.compra.primary} />
          <FilterChip label="Ventas" active={filterType === "venta"} onClick={() => setFilterType("venta")} color={T.venta.primary} />
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <select value={filterClient} onChange={e => setFilterClient(e.target.value)}
            style={{ flex: 1, padding: "10px 12px", borderRadius: 10, background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`, color: C.text, fontSize: 12 }}>
            <option value="all">Todos los clientes</option>
            {allClients.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={filterBank} onChange={e => setFilterBank(e.target.value)}
            style={{ flex: 1, padding: "10px 12px", borderRadius: 10, background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`, color: C.text, fontSize: 12 }}>
            <option value="all">Todos los bancos</option>
            {allBanks.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
        </div>

        {/* Results */}
        <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.16em", marginBottom: 10, fontWeight: 700 }}>
          {results.length} resultado{results.length !== 1 ? "s" : ""}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
          {results.length === 0 ? <Empty label="Sin coincidencias" /> :
            results.map((r, i) => {
              const tc = T[r.op.type];
              return (
                <div key={`${r.op.id}-${r.kind}-${i}`} onClick={() => onDetail(r.op.id)} style={{
                  background: C.card, border: `1px solid ${r.kind === "tt" ? `${C.accent2}40` : C.border}`,
                  borderRadius: 12, padding: "11px 14px",
                  backdropFilter: "blur(16px)",
                  position: "relative", overflow: "hidden", cursor: "pointer",
                  animation: "fadeUp 0.3s ease",
                }}>
                  <div style={{ position: "absolute", top: 0, left: 0, width: 3, bottom: 0, background: `linear-gradient(180deg, ${tc.primary}, ${tc.primary}55)` }} />
                  <div style={{ marginLeft: 8 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 2 }}>
                          <span style={{ fontSize: 14, fontWeight: 700 }}>{r.op.client}</span>
                          <span style={{ fontSize: 8, padding: "2px 6px", borderRadius: 4, background: tc.bg, color: tc.primary, fontWeight: 800, letterSpacing: "0.12em" }}>{tc.label}</span>
                          {r.kind === "tt" && <span style={{ fontSize: 8, padding: "2px 6px", borderRadius: 4, background: C.accent2Dim, color: C.accent2, fontWeight: 800, letterSpacing: "0.12em" }}>TT #{r.idx + 1}</span>}
                        </div>
                        <div style={{ fontSize: 11, color: C.muted }}>{r.op.date.slice(0, 5)} · {r.op.exchange || "—"} · {r.op.bank}</div>
                      </div>
                      <div style={{ textAlign: "right" }}>
                        {r.kind === "tt"
                          ? <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, fontWeight: 700, color: C.accent2 }}>{fmtARS(r.tt.amount)}</div>
                          : <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: tc.primary, fontWeight: 700 }}>{fmtUSDT(r.op.usdtAmount)}</div>}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })
          }
        </div>
      </div>
    </div>
  );
}

function FilterChip({ label, active, onClick, color }) {
  return (
    <button onClick={onClick} style={{
      padding: "8px 14px", borderRadius: 9, whiteSpace: "nowrap",
      background: active ? (color ? `${color}15` : C.accentDim) : "transparent",
      border: `1px solid ${active ? (color || C.accent) : C.border}`,
      color: active ? (color || C.accent) : C.soft,
      fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase",
      fontFamily: "inherit", transition: `all 0.2s ${EASE}`,
    }}>{label}</button>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  BOTTOM NAV
// ═══════════════════════════════════════════════════════════════════
function BottomNav({ tab, onChange, onSettings }) {
  const items = [
    { id: "inicio", label: "Inicio", Icon: ListChecks },
    { id: "ganancias", label: "Ganancias", Icon: TrendingUp },
    { id: "buscar", label: "Buscar", Icon: Search },
  ];
  return (
    <div style={{
      position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)",
      width: "100%", maxWidth: 480, background: C.nav, borderTop: `1px solid ${C.border}`,
      backdropFilter: "blur(28px)", display: "flex", paddingBottom: 22, zIndex: 50,
    }}>
      {items.map(({ id, label, Icon }) => (
        <button key={id} onClick={() => onChange(id)} style={{
          flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 5, paddingTop: 11,
          color: tab === id ? C.accent : C.muted,
          fontSize: 9, fontWeight: tab === id ? 800 : 600,
          letterSpacing: "0.14em", textTransform: "uppercase",
          transition: `color 0.25s ${EASE}`, position: "relative",
        }}>
          {tab === id && <div style={{ position: "absolute", top: 0, left: "50%", transform: "translateX(-50%)", width: 24, height: 2, background: C.accent, borderRadius: "0 0 2px 2px", boxShadow: `0 0 10px ${C.accent}` }} />}
          <div style={{ transform: tab === id ? "scale(1.1)" : "scale(1)", transition: `transform 0.25s ${EASE}` }}>
            <Icon size={18} strokeWidth={tab === id ? 2.2 : 1.8} />
          </div>
          {label}
        </button>
      ))}
      <button onClick={onSettings} style={{
        flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 5, paddingTop: 11,
        color: C.muted, fontSize: 9, fontWeight: 600,
        letterSpacing: "0.14em", textTransform: "uppercase",
      }}>
        <Settings size={18} strokeWidth={1.8} />
        Más
      </button>
    </div>
  );
}

function FloatingCalc({ onClick }) {
  return (
    <button onClick={onClick} style={{
      position: "fixed", right: 18, bottom: 100, zIndex: 60,
      width: 54, height: 54, borderRadius: 18,
      background: `linear-gradient(135deg, ${C.accent}, ${C.accent2}cc)`,
      color: C.bg,
      display: "flex", alignItems: "center", justifyContent: "center",
      boxShadow: `0 10px 32px rgba(232,76,193,0.35), 0 0 0 1px ${C.accent}44`,
      transition: `transform 0.2s ${EASE}`,
    }}
    onTouchStart={e => e.currentTarget.style.transform = "scale(0.92)"}
    onTouchEnd={e => e.currentTarget.style.transform = "scale(1)"}>
      <Calculator size={22} strokeWidth={2.2} />
    </button>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET BASE
// ═══════════════════════════════════════════════════════════════════
function Sheet({ title, subtitle, onClose, children, accent, fullScreen, leftAction }) {
  return (
    <>
      <div onClick={onClose} style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.88)",
        zIndex: 100, backdropFilter: "blur(10px)", animation: "fadeIn 0.3s ease",
      }} />
      <div style={{
        position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)",
        width: "100%", maxWidth: 480,
        background: "rgba(8,4,16,0.99)",
        borderRadius: fullScreen ? "0" : "24px 24px 0 0",
        borderTop: `1px solid ${accent || C.borderStrong}`,
        boxShadow: `0 -24px 80px rgba(0,0,0,0.7), inset 0 1px 0 ${accent || "rgba(180,100,220,0.15)"}`,
        padding: "16px 18px 52px", zIndex: 101,
        maxHeight: fullScreen ? "100vh" : "92vh", height: fullScreen ? "100vh" : "auto",
        overflowY: "auto",
        animation: "slideUp 0.4s " + EASE,
      }}>
        {!fullScreen && <div style={{ width: 36, height: 4, background: "rgba(255,255,255,0.1)", borderRadius: 2, margin: "0 auto 16px" }} />}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: subtitle ? 4 : 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {leftAction}
            <div style={{ fontSize: 20, fontWeight: 800, letterSpacing: "-0.02em" }}>{title}</div>
          </div>
          <button onClick={onClose} style={{
            width: 32, height: 32, borderRadius: 10,
            background: "rgba(255,255,255,0.04)", color: C.soft,
            display: "flex", alignItems: "center", justifyContent: "center",
          }}>
            <X size={18} />
          </button>
        </div>
        {subtitle && <div style={{ fontSize: 12, color: C.muted, marginBottom: 18 }}>{subtitle}</div>}
        {children}
      </div>
    </>
  );
}

function Field({ label, value, onChange, placeholder, mono, type = "text", autoFocus }) {
  return (
    <div>
      <div style={{ fontSize: 10, color: C.muted, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.14em", fontWeight: 700 }}>{label}</div>
      <input
        value={value} onChange={e => onChange(e.target.value)}
        placeholder={placeholder} autoFocus={autoFocus}
        inputMode={mono ? "decimal" : type === "tel" ? "tel" : "text"}
        style={{
          width: "100%", padding: "12px 14px", borderRadius: 11,
          background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`,
          color: C.text, fontSize: 15,
          fontFamily: mono ? "'JetBrains Mono', monospace" : "inherit",
          transition: `border-color 0.2s ${EASE}`,
        }}
        onFocus={e => e.target.style.borderColor = C.accent + "60"}
        onBlur={e => e.target.style.borderColor = C.border}
      />
    </div>
  );
}

function Empty({ label, cta }) {
  return (
    <div style={{ textAlign: "center", padding: "60px 0", color: C.muted, animation: "fadeUp 0.4s ease" }}>
      <div style={{ width: 56, height: 56, margin: "0 auto 14px", borderRadius: 16, background: C.accentDim, border: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <AtomLogo size={28} />
      </div>
      <div style={{ fontSize: 12, fontWeight: 600, letterSpacing: "0.12em", textTransform: "uppercase", marginBottom: 4 }}>{label}</div>
      {cta && <div style={{ fontSize: 11, color: C.soft, marginTop: 6 }}>{cta}</div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · NEW OP
// ═══════════════════════════════════════════════════════════════════
function NewOpSheet({ onClose, onSave, clientHistory }) {
  const [f, setF] = useState({ type: "compra", client: "", usdtAmount: "", tcClient: "", tcOtc: "", bank: "Ripio", exchange: "", notes: "" });
  const [showSuggest, setShowSuggest] = useState(false);
  const set = (k, v) => setF(p => ({ ...p, [k]: v }));
  const tc = T[f.type];

  const valid = f.client && f.usdtAmount && f.tcClient && f.bank;

  const clientSuggestions = useMemo(() => {
    if (!f.client) return [];
    const q = f.client.toLowerCase();
    return Object.keys(clientHistory).filter(c => c.toLowerCase().startsWith(q) && c.toLowerCase() !== q).slice(0, 5);
  }, [f.client, clientHistory]);

  const pickClient = (name) => {
    const hist = clientHistory[name];
    setF(p => ({
      ...p, client: name,
      bank: p.bank || hist?.bank || "Ripio",
      exchange: p.exchange || hist?.exchange || "",
      type: hist?.type || p.type,
    }));
    setShowSuggest(false);
  };

  const preview = f.usdtAmount && f.tcClient ? {
    total: parseNum(f.usdtAmount) * parseNum(f.tcClient),
    profit: f.tcOtc ? (f.type === "compra" ? parseNum(f.tcClient) - parseNum(f.tcOtc) : parseNum(f.tcOtc) - parseNum(f.tcClient)) * parseNum(f.usdtAmount) : null,
  } : null;

  const submit = () => {
    if (!valid) return;
    onSave({
      type: f.type, client: f.client.trim(), date: todayStr(),
      usdtAmount: parseNum(f.usdtAmount), tcClient: parseNum(f.tcClient),
      tcOtc: f.tcOtc ? parseNum(f.tcOtc) : null,
      bank: f.bank.trim(), exchange: f.exchange.trim(), notes: f.notes.trim(),
    });
  };

  return (
    <Sheet title="Nueva operación" onClose={onClose} accent={tc.primary}>
      {/* Type */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 18 }}>
        {["compra", "venta"].map(type => {
          const tt = T[type]; const active = f.type === type;
          return (
            <button key={type} onClick={() => set("type", type)} style={{
              padding: "13px", borderRadius: 12, fontSize: 12, fontWeight: 800,
              letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: "inherit",
              background: active ? tt.dim : "transparent",
              border: `1px solid ${active ? tt.border : C.border}`,
              color: active ? tt.primary : C.muted,
              boxShadow: active ? tt.glow : "none",
              transition: `all 0.25s ${EASE}`,
            }}>{tt.label}</button>
          );
        })}
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {/* Client with suggestions */}
        <div style={{ position: "relative" }}>
          <Field label="Cliente" value={f.client}
            onChange={v => { set("client", v); setShowSuggest(true); }}
            placeholder="Nombre del cliente" autoFocus />
          {showSuggest && clientSuggestions.length > 0 && (
            <div style={{
              position: "absolute", top: "100%", left: 0, right: 0, marginTop: 4,
              background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
              zIndex: 10, overflow: "hidden", backdropFilter: "blur(20px)",
              boxShadow: "0 10px 30px rgba(0,0,0,0.5)",
            }}>
              {clientSuggestions.map(c => (
                <button key={c} onClick={() => pickClient(c)} style={{
                  width: "100%", padding: "10px 14px", textAlign: "left",
                  color: C.text, fontSize: 13, fontFamily: "inherit",
                  borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center",
                }}>
                  <span>{c}</span>
                  <span style={{ fontSize: 10, color: C.muted }}>{clientHistory[c]?.bank} · {clientHistory[c]?.exchange || "—"}</span>
                </button>
              ))}
            </div>
          )}
        </div>

        <Field label="USDT" value={f.usdtAmount} onChange={v => set("usdtAmount", v)} placeholder="0,00" mono />

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="TC Cliente" value={f.tcClient} onChange={v => set("tcClient", v)} placeholder="0" mono />
          <Field label="TC OTC (opcional)" value={f.tcOtc} onChange={v => set("tcOtc", v)} placeholder="0" mono />
        </div>

        <div>
          <div style={{ fontSize: 10, color: C.muted, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.14em", fontWeight: 700 }}>Banco</div>
          <select value={f.bank} onChange={e => set("bank", e.target.value)}
            style={{ width: "100%", padding: "12px 14px", borderRadius: 11, background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`, color: C.text, fontSize: 15 }}>
            {BANKS.map(b => <option key={b.id} value={b.id}>{b.label}</option>)}
          </select>
        </div>

        <Field label="Exchange" value={f.exchange} onChange={v => set("exchange", v)} placeholder="Ej: Bybit, Cocos, Binance…" />

        <div>
          <div style={{ fontSize: 10, color: C.muted, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.14em", fontWeight: 700 }}>Notas (privadas)</div>
          <textarea value={f.notes} onChange={e => set("notes", e.target.value)}
            placeholder="Anotaciones internas, no se muestran en presentación al cliente…"
            style={{
              width: "100%", padding: "11px 14px", borderRadius: 11, minHeight: 60, resize: "vertical",
              background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`,
              color: C.text, fontSize: 13, fontFamily: "inherit",
            }} />
        </div>

        {/* Preview */}
        {preview && (
          <div style={{ background: tc.dim, border: `1px solid ${tc.border}`, borderRadius: 12, padding: "12px 14px", animation: "fadeUp 0.3s ease" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <span style={{ fontSize: 11, color: C.soft, textTransform: "uppercase", letterSpacing: "0.1em" }}>Total ARS</span>
              <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, fontWeight: 700, color: C.text }}>{fmtARS(preview.total)}</span>
            </div>
            {preview.profit !== null && (
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 11, color: C.soft, textTransform: "uppercase", letterSpacing: "0.1em" }}>Ganancia est.</span>
                <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, fontWeight: 700, color: preview.profit >= 0 ? C.positive : C.negative }}>{fmtARS(preview.profit)}</span>
              </div>
            )}
          </div>
        )}

        <button onClick={submit} disabled={!valid} style={{
          width: "100%", padding: 14, borderRadius: 12, fontFamily: "inherit",
          background: valid ? tc.dim : "transparent",
          border: `1px solid ${valid ? tc.border : C.border}`,
          color: valid ? tc.primary : C.muted,
          fontSize: 12, fontWeight: 800, letterSpacing: "0.14em", textTransform: "uppercase",
          boxShadow: valid ? tc.glow : "none", transition: `all 0.25s ${EASE}`,
        }}>Guardar operación</button>
      </div>
    </Sheet>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · OP DETAIL (edit any field, share, present, duplicate)
// ═══════════════════════════════════════════════════════════════════
function OpDetailSheet({ op, onClose, onAddTTs, onDelTT, onDelOp, onUpdate, onDuplicate, onPresent, notify }) {
  const [waMsg, setWaMsg] = useState("");
  const [parsing, setParsing] = useState(false);
  const [parseResult, setParseResult] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [confirmDel, setConfirmDel] = useState(false);
  const [editing, setEditing] = useState(null); // field name being edited
  const [editValue, setEditValue] = useState("");

  const tc = T[op.type];
  const total = op.usdtAmount * op.tcClient;
  const sent = op.tts.reduce((s, t) => s + t.amount, 0);
  const remaining = total - sent;
  const done = Math.abs(remaining) < 1;
  const overshoot = remaining < 0 && !done;
  const pct = total > 0 ? Math.min(100, Math.max(0, (sent / total) * 100)) : 0;
  const profit = op.tcOtc ? (op.type === "compra" ? op.tcClient - op.tcOtc : op.tcOtc - op.tcClient) * op.usdtAmount : null;

  const handleParse = async () => {
    if (!waMsg.trim()) return;
    setParsing(true);
    try {
      const result = await parseWhatsApp(waMsg);
      // Check for duplicates
      const existingAmounts = new Set(op.tts.map(t => Math.round(t.amount)));
      const duplicates = result.amounts.filter(a => existingAmounts.has(Math.round(a)));
      if (duplicates.length > 0) {
        result.warnings = [...(result.warnings || []), `⚠ Posible duplicado: ${duplicates.map(fmtARS).join(", ")} ya está cargado en esta op`];
      }
      setParseResult(result);
    } catch { setParseResult({ amounts: [], warnings: ["Error al procesar"] }); }
    setParsing(false);
  };
  const confirmAdd = () => {
    if (!parseResult?.amounts?.length) return;
    onAddTTs(op.id, parseResult.amounts);
    setWaMsg(""); setParseResult(null); setShowAdd(false);
  };

  const startEdit = (field, value) => { setEditing(field); setEditValue(String(value ?? "")); };
  const fieldLabels = { usdtAmount: "USDT", tcClient: "TC cliente", tcOtc: "TC OTC", bank: "Banco", exchange: "Exchange" };
  const saveEdit = () => {
    if (!editing) return;
    let v = editValue;
    if (["usdtAmount", "tcClient", "tcOtc"].includes(editing)) v = editing === "tcOtc" && !editValue ? null : parseNum(editValue);
    onUpdate(op.id, { [editing]: v });
    setEditing(null);
    notify(`✓ ${fieldLabels[editing] || editing} actualizado`);
  };

  const copyProgress = () => {
    const lines = [];
    lines.push(`PLEXO · Operación ${tc.label}`);
    lines.push(`─────────────────────`);
    lines.push(`Cliente: ${op.client}`);
    lines.push(`Fecha: ${op.date}`);
    lines.push(`USDT: ${fmtUSDT(op.usdtAmount)}`);
    lines.push(`TC: ${op.tcClient.toLocaleString("es-AR")}`);
    lines.push(`Total: ${fmtARS(total)}`);
    lines.push(``);
    lines.push(`Enviado: ${fmtARS(sent)} en ${op.tts.length} TT${op.tts.length !== 1 ? "s" : ""}`);
    lines.push(``);
    if (done) {
      lines.push(`✓ Operación saldada`);
    } else if (overshoot) {
      lines.push(`ARS a favor Gonza: ${op.client} debe ${fmtARS(-remaining)}`);
    } else {
      lines.push(`ARS a favor ${op.client}: ${fmtARS(remaining)}`);
    }
    const text = lines.join("\n");
    navigator.clipboard.writeText(text).then(() => notify("✓ Copiado al portapapeles")).catch(() => notify("Error al copiar", { kind: "error" }));
  };

  const exportClient = () => {
    exportClientXLSX(op);
    notify("✓ Excel del cliente descargado");
  };

  return (
    <Sheet title={op.client} onClose={onClose} accent={tc.primary}>
      {/* Meta */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <span style={{ fontSize: 10, fontWeight: 800, padding: "3px 9px", borderRadius: 5, background: tc.bg, color: tc.primary, letterSpacing: "0.14em" }}>{tc.label}</span>
        <span style={{ fontSize: 11, color: C.muted }}>{op.date} · {op.exchange || "—"} · {op.bank}</span>
      </div>

      {/* Info grid editable */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`, borderRadius: 13, padding: "13px 14px", marginBottom: 12 }}>
        <EditableField label="USDT" field="usdtAmount" value={fmtUSDT(op.usdtAmount)} mono color={tc.primary}
          editing={editing} editValue={editValue} setEditValue={setEditValue} onStart={() => startEdit("usdtAmount", op.usdtAmount)} onSave={saveEdit} onCancel={() => setEditing(null)} />
        <InfoItem label="Total ARS" value={fmtARS(total)} mono />
        <EditableField label="TC Cliente" field="tcClient" value={op.tcClient.toLocaleString("es-AR")} mono
          editing={editing} editValue={editValue} setEditValue={setEditValue} onStart={() => startEdit("tcClient", op.tcClient)} onSave={saveEdit} onCancel={() => setEditing(null)} />
        <EditableField label="TC OTC" field="tcOtc" value={op.tcOtc ? op.tcOtc.toLocaleString("es-AR") : "—"} mono color={op.tcOtc ? C.text : C.warn}
          editing={editing} editValue={editValue} setEditValue={setEditValue} onStart={() => startEdit("tcOtc", op.tcOtc)} onSave={saveEdit} onCancel={() => setEditing(null)} />
        <EditableField label="Banco" field="bank" value={op.bank}
          editing={editing} editValue={editValue} setEditValue={setEditValue} onStart={() => startEdit("bank", op.bank)} onSave={saveEdit} onCancel={() => setEditing(null)}
          options={BANKS.map(b => b.id)} />
        <EditableField label="Exchange" field="exchange" value={op.exchange || "—"}
          editing={editing} editValue={editValue} setEditValue={setEditValue} onStart={() => startEdit("exchange", op.exchange)} onSave={saveEdit} onCancel={() => setEditing(null)} />
      </div>

      {/* Profit */}
      {profit !== null && (
        <div style={{ background: profit >= 0 ? C.positiveDim : C.negativeDim, border: `1px solid ${profit >= 0 ? C.positiveBorder : C.negativeBorder}`, borderRadius: 12, padding: "11px 14px", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 11, color: C.soft, textTransform: "uppercase", letterSpacing: "0.1em" }}>Ganancia</span>
          <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 16, fontWeight: 700, color: profit >= 0 ? C.positive : C.negative }}>{fmtARS(profit)}</span>
        </div>
      )}

      {/* Progress */}
      <div style={{ background: done ? tc.dim : overshoot ? C.warnDim : C.accentDim, border: `1px solid ${done ? tc.border : overshoot ? `${C.warn}44` : C.accent + "33"}`, borderRadius: 13, padding: "13px 14px", marginBottom: 12, boxShadow: done ? tc.glow : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 10 }}>
          <span style={{ color: C.soft, textTransform: "uppercase", letterSpacing: "0.1em", fontSize: 10, fontWeight: 700 }}>Progreso</span>
          <span style={{ fontFamily: "'JetBrains Mono', monospace", color: done ? tc.primary : overshoot ? C.warn : C.accent }}>
            {fmtARS(sent)} / {fmtARS(total)}
          </span>
        </div>
        <div style={{ height: 6, background: "rgba(255,255,255,0.05)", borderRadius: 3, overflow: "hidden" }}>
          <div style={{ height: "100%", width: `${pct}%`, background: `linear-gradient(90deg, ${done ? tc.primary : C.accent}88, ${done ? tc.primary : C.accent})`, borderRadius: 3, transition: `width 0.5s ${EASE}` }} />
        </div>
        <div style={{ marginTop: 9, textAlign: "right", fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: done ? tc.primary : overshoot ? C.warn : C.accent2, fontWeight: 700 }}>
          {done ? "✓ Saldado" : overshoot ? `${op.client} debe ${fmtARS(-remaining)}` : `A favor ${op.client}: ${fmtARS(remaining)}`}
        </div>
      </div>

      {/* Actions row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6, marginBottom: 14 }}>
        <ActionBtn icon={Copy} label="Copiar" onClick={copyProgress} />
        <ActionBtn icon={Eye} label="Mostrar" onClick={onPresent} />
        <ActionBtn icon={FileText} label="Excel" onClick={exportClient} />
        <ActionBtn icon={Copy} label="Duplicar" onClick={() => onDuplicate(op.id)} />
      </div>

      {/* Notes */}
      {op.notes && (
        <div style={{ background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`, borderRadius: 11, padding: "10px 13px", marginBottom: 12 }}>
          <div style={{ fontSize: 9, color: C.muted, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 4, fontWeight: 700, display: "flex", alignItems: "center", gap: 6 }}>
            <ScrollText size={11} /> Nota privada
          </div>
          <div style={{ fontSize: 13, color: C.textDim, whiteSpace: "pre-wrap", lineHeight: 1.4 }}>{op.notes}</div>
        </div>
      )}
      <button onClick={() => { const n = prompt("Nota privada (vacío para borrar):", op.notes || ""); if (n !== null) { onUpdate(op.id, { notes: n }); notify(n ? "Nota guardada" : "Nota borrada"); } }} style={{ width: "100%", padding: "10px", borderRadius: 10, border: `1px dashed ${C.border}`, color: C.soft, fontSize: 10, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", marginBottom: 14, fontFamily: "inherit" }}>
        <Pencil size={11} style={{ verticalAlign: "middle", marginRight: 6 }} />
        {op.notes ? "Editar nota" : "Agregar nota"}
      </button>

      {/* TTs */}
      {op.tts.length > 0 && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.18em", marginBottom: 8, display: "flex", justifyContent: "space-between", fontWeight: 700 }}>
            <span>Transferencias · {op.tts.length}</span>
            <span style={{ fontFamily: "'JetBrains Mono', monospace", color: C.text }}>{fmtARS(sent)}</span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
            {op.tts.map((tt, i) => (
              <div key={tt.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`, borderRadius: 10, padding: "9px 14px" }}>
                <span style={{ fontSize: 11, color: C.muted, fontFamily: "'JetBrains Mono', monospace" }}>TT {String(i + 1).padStart(2, "0")}</span>
                <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14 }}>{fmtARS(tt.amount)}</span>
                <button onClick={() => onDelTT(op.id, tt.id)} style={{ color: "rgba(251,113,133,0.55)", padding: "0 2px" }}>
                  <X size={16} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Add TT */}
      {!showAdd
        ? <button onClick={() => setShowAdd(true)} style={{ width: "100%", padding: 12, borderRadius: 12, border: `1px dashed ${C.accent}66`, color: C.accent, fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase", marginBottom: 14, fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            <Plus size={14} /> Agregar TT desde WhatsApp
          </button>
        : <AddTTBox waMsg={waMsg} setWaMsg={setWaMsg} parsing={parsing} parseResult={parseResult} handleParse={handleParse} confirmAdd={confirmAdd} onCancel={() => { setShowAdd(false); setWaMsg(""); setParseResult(null); }} />
      }

      {/* Delete */}
      {!confirmDel
        ? <button onClick={() => setConfirmDel(true)} style={{ width: "100%", padding: 12, borderRadius: 12, border: `1px solid ${C.negative}33`, color: C.negative + "aa", fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase", fontFamily: "inherit" }}>
            Eliminar operación
          </button>
        : <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <button onClick={() => setConfirmDel(false)} style={{ padding: 12, borderRadius: 12, border: `1px solid ${C.border}`, color: C.soft, fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "inherit" }}>Cancelar</button>
            <button onClick={() => onDelOp(op.id)} style={{ padding: 12, borderRadius: 12, background: `${C.negative}1a`, border: `1px solid ${C.negative}40`, color: C.negative, fontSize: 11, fontWeight: 800, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "inherit" }}>Eliminar</button>
          </div>
      }
    </Sheet>
  );
}

function ActionBtn({ icon: Icon, label, onClick }) {
  return (
    <button onClick={onClick} style={{
      padding: "10px 6px", borderRadius: 11,
      background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`,
      color: C.text, fontSize: 9, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase",
      fontFamily: "inherit", display: "flex", flexDirection: "column", alignItems: "center", gap: 4,
      transition: `all 0.2s ${EASE}`,
    }}>
      <Icon size={14} color={C.accent2} strokeWidth={1.8} />
      {label}
    </button>
  );
}

function InfoItem({ label, value, mono, color }) {
  return (
    <div>
      <div style={{ fontSize: 9, color: C.muted, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 4, fontWeight: 700 }}>{label}</div>
      <div style={{ fontSize: 13, fontFamily: mono ? "'JetBrains Mono', monospace" : "inherit", color: color || C.text, fontWeight: 600 }}>{value}</div>
    </div>
  );
}
function EditableField({ label, field, value, mono, color, editing, editValue, setEditValue, onStart, onSave, onCancel, options }) {
  const isEditing = editing === field;
  return (
    <div>
      <div style={{ fontSize: 9, color: C.muted, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 4, fontWeight: 700 }}>{label}</div>
      {isEditing ? (
        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
          {options ? (
            <select value={editValue} onChange={e => setEditValue(e.target.value)}
              style={{ flex: 1, background: "rgba(255,255,255,0.05)", border: `1px solid ${C.accent}66`, borderRadius: 7, padding: "4px 8px", color: C.text, fontSize: 13 }}>
              {options.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          ) : (
            <input autoFocus value={editValue} onChange={e => setEditValue(e.target.value)}
              inputMode={mono ? "decimal" : "text"}
              style={{ flex: 1, background: "rgba(255,255,255,0.05)", border: `1px solid ${C.accent}66`, borderRadius: 7, padding: "4px 8px", color: C.text, fontSize: 13, fontFamily: mono ? "'JetBrains Mono', monospace" : "inherit" }} />
          )}
          <button onClick={onSave} style={{ color: C.positive, padding: "0 2px" }}><Check size={16} /></button>
          <button onClick={onCancel} style={{ color: C.muted, padding: "0 2px" }}><X size={14} /></button>
        </div>
      ) : (
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 13, fontFamily: mono ? "'JetBrains Mono', monospace" : "inherit", color: color || C.text, fontWeight: 600 }}>{value}</span>
          <button onClick={onStart} style={{ color: C.muted, padding: 0 }}><Pencil size={11} /></button>
        </div>
      )}
    </div>
  );
}

function AddTTBox({ waMsg, setWaMsg, parsing, parseResult, handleParse, confirmAdd, onCancel }) {
  return (
    <div style={{ background: C.accentDim, border: `1px solid ${C.accent}33`, borderRadius: 13, padding: 14, marginBottom: 14, animation: "fadeUp 0.3s ease" }}>
      <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 8, fontWeight: 700 }}>
        Pegá el mensaje de WhatsApp
      </div>
      <textarea value={waMsg} onChange={e => setWaMsg(e.target.value)}
        placeholder="Pegá el texto del mensaje (o varios)…"
        style={{
          width: "100%", minHeight: 80, background: "rgba(0,0,0,0.4)",
          border: `1px solid ${C.border}`, borderRadius: 10, padding: 12,
          color: C.text, fontSize: 13, resize: "vertical", fontFamily: "inherit",
        }} />
      {parseResult === null
        ? <button onClick={handleParse} disabled={parsing || !waMsg.trim()} style={{
            width: "100%", padding: 11, borderRadius: 10, marginTop: 10,
            background: parsing || !waMsg.trim() ? "transparent" : C.accentDim,
            border: `1px solid ${parsing || !waMsg.trim() ? C.border : C.accent + "40"}`,
            color: parsing || !waMsg.trim() ? C.muted : C.accent,
            fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase",
            fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
          }}>
            {parsing ? <><Loader2 size={13} className="spin" style={{ animation: "spin 1s linear infinite" }} /> Analizando…</> : <><Sparkles size={13} /> Extraer montos</>}
          </button>
        : <div style={{ marginTop: 12 }}>
            {parseResult.warnings?.length > 0 && (
              <div style={{ marginBottom: 10 }}>
                {parseResult.warnings.map((w, i) => (
                  <div key={i} style={{ display: "flex", gap: 8, fontSize: 11, color: C.warn, background: C.warnDim, border: `1px solid ${C.warn}33`, padding: "8px 10px", borderRadius: 8, marginBottom: 4 }}>
                    <AlertCircle size={13} style={{ flexShrink: 0, marginTop: 1 }} />
                    <span>{w}</span>
                  </div>
                ))}
              </div>
            )}
            {parseResult.amounts?.length === 0
              ? <div style={{ color: C.negative, fontSize: 12, padding: "8px 0" }}>No se detectaron montos</div>
              : <>
                  <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 8, fontWeight: 700 }}>Detectados · {parseResult.amounts.length}</div>
                  {parseResult.amounts.map((a, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "'JetBrains Mono', monospace", fontSize: 14, color: C.positive, marginBottom: 4 }}>
                      <Check size={14} /> {fmtARS(a)}
                    </div>
                  ))}
                </>
            }
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 12 }}>
              <button onClick={onCancel} style={{ padding: 11, borderRadius: 10, border: `1px solid ${C.border}`, color: C.soft, fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "inherit" }}>Cancelar</button>
              <button onClick={confirmAdd} disabled={!parseResult.amounts?.length} style={{
                padding: 11, borderRadius: 10,
                background: parseResult.amounts?.length ? C.positiveDim : "transparent",
                border: `1px solid ${parseResult.amounts?.length ? C.positiveBorder : C.border}`,
                color: parseResult.amounts?.length ? C.positive : C.muted,
                fontSize: 11, fontWeight: 800, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "inherit",
              }}>Confirmar</button>
            </div>
          </div>
      }
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · QUICK ADD TT (desde home)
// ═══════════════════════════════════════════════════════════════════
function QuickTTSheet({ op, onClose, onAddTTs, notify }) {
  const [waMsg, setWaMsg] = useState("");
  const [parsing, setParsing] = useState(false);
  const [parseResult, setParseResult] = useState(null);
  const tc = T[op.type];

  const handleParse = async () => {
    if (!waMsg.trim()) return;
    setParsing(true);
    try {
      const result = await parseWhatsApp(waMsg);
      const existingAmounts = new Set(op.tts.map(t => Math.round(t.amount)));
      const duplicates = result.amounts.filter(a => existingAmounts.has(Math.round(a)));
      if (duplicates.length > 0) {
        result.warnings = [...(result.warnings || []), `⚠ Posible duplicado: ${duplicates.map(fmtARS).join(", ")}`];
      }
      setParseResult(result);
    } catch { setParseResult({ amounts: [], warnings: ["Error"] }); }
    setParsing(false);
  };

  const confirmAdd = () => {
    if (!parseResult?.amounts?.length) return;
    onAddTTs(op.id, parseResult.amounts);
    onClose();
  };

  return (
    <Sheet title={`+ TT a ${op.client}`} subtitle={`${tc.label} · ${fmtUSDT(op.usdtAmount)} · TC ${op.tcClient.toLocaleString("es-AR")}`} onClose={onClose} accent={tc.primary}>
      <AddTTBox waMsg={waMsg} setWaMsg={setWaMsg} parsing={parsing} parseResult={parseResult} handleParse={handleParse} confirmAdd={confirmAdd} onCancel={onClose} />
    </Sheet>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · PRESENTACIÓN PARA CLIENTE (fullscreen, sin info privada)
// ═══════════════════════════════════════════════════════════════════
function PresentSheet({ op, onClose, notify }) {
  const tc = T[op.type];
  const total = op.usdtAmount * op.tcClient;
  const sent = op.tts.reduce((s, t) => s + t.amount, 0);
  const remaining = total - sent;
  const done = Math.abs(remaining) < 1;
  const overshoot = remaining < 0 && !done;
  const pct = total > 0 ? Math.min(100, Math.max(0, (sent / total) * 100)) : 0;

  return (
    <>
      <div style={{ position: "fixed", inset: 0, background: C.bg, zIndex: 200, animation: "fadeIn 0.3s ease", overflowY: "auto" }}>
        <BgGrid />
        <div style={{ position: "relative", zIndex: 1, padding: "20px 20px 40px", maxWidth: 480, margin: "0 auto" }}>
          {/* Header */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 36 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <AtomLogo size={26} />
              <span style={{ fontSize: 11, color: C.muted, letterSpacing: "0.3em", textTransform: "uppercase", fontWeight: 800 }}>Plexo</span>
            </div>
            <button onClick={onClose} style={{ width: 36, height: 36, borderRadius: 11, background: "rgba(255,255,255,0.06)", color: C.text, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <X size={18} />
            </button>
          </div>

          {/* Client + type badge */}
          <div style={{ textAlign: "center", marginBottom: 28 }}>
            <span style={{ fontSize: 10, fontWeight: 800, padding: "4px 12px", borderRadius: 6, background: tc.bg, color: tc.primary, letterSpacing: "0.18em" }}>{tc.label}</span>
            <h1 style={{ fontSize: 36, fontWeight: 800, letterSpacing: "-0.03em", margin: "16px 0 6px", lineHeight: 1 }}>{op.client}</h1>
            <div style={{ fontSize: 11, color: C.muted, letterSpacing: "0.16em", textTransform: "uppercase", fontWeight: 600 }}>{op.date}</div>
          </div>

          {/* Stats grid */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 22 }}>
            <BigStat label="USDT" value={fmtUSDT(op.usdtAmount)} color={tc.primary} />
            <BigStat label="Tipo de cambio" value={op.tcClient.toLocaleString("es-AR")} />
            <BigStat label="Total ARS" value={fmtARS(total)} fullWidth />
          </div>

          {/* Progress hero */}
          <div style={{
            background: done ? C.positiveDim : overshoot ? C.negativeDim : C.card,
            border: `1px solid ${done ? C.positiveBorder : overshoot ? C.negativeBorder : C.border}`,
            borderRadius: 18, padding: 22, marginBottom: 22,
            backdropFilter: "blur(20px)",
            boxShadow: done ? `0 0 40px ${C.positive}20` : overshoot ? `0 0 40px ${C.negative}20` : "none",
            textAlign: "center",
          }}>
            <div style={{ fontSize: 10, color: done ? C.positive : overshoot ? C.negative : C.muted, textTransform: "uppercase", letterSpacing: "0.22em", marginBottom: 12, fontWeight: 800 }}>
              {done ? "✓ Saldado" : overshoot ? "Sobrante" : "Pendiente"}
            </div>
            <div style={{ fontSize: 34, fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, letterSpacing: "-0.02em", marginBottom: 18, color: done ? C.positive : overshoot ? C.negative : C.text, lineHeight: 1 }}>
              {done ? fmtARS(total) : overshoot ? fmtARS(-remaining) : fmtARS(remaining)}
            </div>
            <div style={{ height: 6, background: "rgba(255,255,255,0.05)", borderRadius: 3, overflow: "hidden", marginBottom: 14 }}>
              <div style={{ height: "100%", width: `${pct}%`, background: done ? `linear-gradient(90deg, ${C.positive}, ${C.positive})` : overshoot ? `linear-gradient(90deg, ${C.negative}, ${C.negative})` : `linear-gradient(90deg, ${C.accent}, ${C.accent2})`, borderRadius: 3, transition: `width 1s ${EASE}` }} />
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11 }}>
              <div>
                <div style={{ color: C.muted, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 3, fontWeight: 700 }}>Enviado</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, fontWeight: 600 }}>{fmtARS(sent)}</div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div style={{ color: C.muted, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.14em", marginBottom: 3, fontWeight: 700 }}>Esperado</div>
                <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, fontWeight: 600 }}>{fmtARS(total)}</div>
              </div>
            </div>
          </div>

          {/* Status message elegant */}
          {!done && (
            <div style={{
              background: overshoot ? C.negativeDim : C.positiveDim,
              border: `1px solid ${overshoot ? C.negativeBorder : C.positiveBorder}`,
              borderRadius: 13, padding: "14px 18px", marginBottom: 22,
              textAlign: "center",
            }}>
              <div style={{ fontSize: 10, color: overshoot ? C.negative : C.positive, textTransform: "uppercase", letterSpacing: "0.18em", marginBottom: 4, fontWeight: 800 }}>
                {overshoot ? `${op.client} debe` : `ARS a favor ${op.client}`}
              </div>
              <div style={{ fontSize: 22, fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, color: overshoot ? C.negative : C.positive }}>
                {fmtARS(Math.abs(remaining))}
              </div>
            </div>
          )}

          {/* TTs */}
          {op.tts.length > 0 && (
            <div>
              <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.18em", marginBottom: 10, fontWeight: 700, display: "flex", justifyContent: "space-between" }}>
                <span>Transferencias · {op.tts.length}</span>
                <span style={{ fontFamily: "'JetBrains Mono', monospace", color: C.text }}>{fmtARS(sent)}</span>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                {op.tts.map((tt, i) => (
                  <div key={tt.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.card, border: `1px solid ${C.border}`, borderRadius: 11, padding: "11px 14px" }}>
                    <span style={{ fontSize: 11, color: C.muted, fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.06em" }}>TT {String(i + 1).padStart(2, "0")}</span>
                    <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, fontWeight: 600 }}>{fmtARS(tt.amount)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </>
  );
}

function BigStat({ label, value, color, fullWidth }) {
  return (
    <div style={{
      gridColumn: fullWidth ? "1 / -1" : "auto",
      background: C.card, border: `1px solid ${C.border}`,
      borderRadius: 14, padding: "13px 16px",
      backdropFilter: "blur(16px)",
    }}>
      <div style={{ fontSize: 9, color: C.muted, textTransform: "uppercase", letterSpacing: "0.16em", marginBottom: 5, fontWeight: 700 }}>{label}</div>
      <div style={{ fontSize: fullWidth ? 22 : 16, fontFamily: "'JetBrains Mono', monospace", fontWeight: 700, color: color || C.text, letterSpacing: "-0.01em" }}>{value}</div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · CALCULADORA DE COMISIONES
// ═══════════════════════════════════════════════════════════════════
function CalcSheet({ onClose }) {
  const [type, setType] = useState("compra");
  const [tcLow, setTcLow] = useState("");
  const [tcHigh, setTcHigh] = useState("");
  const [usdtInput, setUsdtInput] = useState("");
  const [bankOut, setBankOut] = useState("Ripio");
  const [bankIn, setBankIn] = useState("Ripio");

  const tc = T[type];
  const lo = parseNum(tcLow), hi = parseNum(tcHigh), usdt = parseNum(usdtInput);

  const bOut = BANKS.find(b => b.id === bankOut)?.comm || 0;
  const bIn = BANKS.find(b => b.id === bankIn)?.comm || 0;

  const valid = lo > 0 && hi > 0 && usdt > 0;

  const calc = useMemo(() => {
    if (!valid) return null;
    const tcLowN = type === "compra" ? lo : hi;  // TC al que comprás USDT
    const tcHighN = type === "compra" ? hi : lo; // TC al que vendés USDT
    const grossSpread = ((hi - lo) / lo) * 100;
    const arsGross = usdt * (hi - lo);
    const binanceFee = usdt * tcLowN * BINANCE_FEE;
    const bankOutFee = usdt * tcLowN * bOut;  // sale ARS
    const bankInFee = usdt * tcHighN * bIn;   // entra ARS
    const totalFee = binanceFee + bankOutFee + bankInFee;
    const arsNet = arsGross - totalFee;
    const netSpread = ((arsNet / (usdt * tcLowN)) * 100);
    return { grossSpread, arsGross, binanceFee, bankOutFee, bankInFee, totalFee, arsNet, netSpread };
  }, [valid, lo, hi, usdt, bOut, bIn, type]);

  return (
    <Sheet title="Calculadora" subtitle="Comisiones + Binance 0,16%" onClose={onClose} accent={C.accent2}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
        {["compra", "venta"].map(t => {
          const tt = T[t]; const active = type === t;
          return (
            <button key={t} onClick={() => setType(t)} style={{
              padding: "12px", borderRadius: 11, fontSize: 11, fontWeight: 800,
              letterSpacing: "0.14em", textTransform: "uppercase",
              background: active ? tt.dim : "transparent",
              border: `1px solid ${active ? tt.border : C.border}`,
              color: active ? tt.primary : C.muted, fontFamily: "inherit",
              transition: `all 0.25s ${EASE}`,
            }}>{tt.label}</button>
          );
        })}
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 11 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="TC bajo" value={tcLow} onChange={setTcLow} placeholder="0" mono />
          <Field label="TC alto" value={tcHigh} onChange={setTcHigh} placeholder="0" mono />
        </div>
        <Field label="USDT" value={usdtInput} onChange={setUsdtInput} placeholder="0,00" mono />

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <BankSelect label="Banco envío" value={bankOut} onChange={setBankOut} />
          <BankSelect label="Banco recibo" value={bankIn} onChange={setBankIn} />
        </div>

        {/* Result */}
        {calc && (
          <div style={{
            background: calc.arsNet >= 0 ? C.positiveDim : C.negativeDim,
            border: `1px solid ${calc.arsNet >= 0 ? C.positiveBorder : C.negativeBorder}`,
            borderRadius: 14, padding: "16px 18px", marginTop: 8,
            backdropFilter: "blur(16px)", animation: "fadeUp 0.3s ease",
          }}>
            <div style={{ fontSize: 10, color: calc.arsNet >= 0 ? C.positive : C.negative, textTransform: "uppercase", letterSpacing: "0.18em", marginBottom: 8, fontWeight: 800 }}>
              Ganancia neta
            </div>
            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 26, fontWeight: 700, color: calc.arsNet >= 0 ? C.positive : C.negative, marginBottom: 4 }}>
              {fmtARS(calc.arsNet)}
            </div>
            <div style={{ fontSize: 11, color: C.soft, marginBottom: 14 }}>
              Spread neto: <span style={{ fontFamily: "'JetBrains Mono', monospace", color: calc.netSpread >= 0 ? C.positive : C.negative }}>{fmtPct(calc.netSpread)}</span>
              {" · "}
              Bruto: <span style={{ fontFamily: "'JetBrains Mono', monospace", color: C.textDim }}>{fmtPct(calc.grossSpread)}</span>
            </div>
            <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12, fontSize: 12 }}>
              <CalcRow label="Spread bruto ARS" value={fmtARS(calc.arsGross)} />
              <CalcRow label="− Binance 0,16%" value={`−${fmtARS(calc.binanceFee)}`} dim />
              {bOut > 0 && <CalcRow label={`− ${bankOut} ${(bOut * 100).toFixed(2).replace(".", ",")}%`} value={`−${fmtARS(calc.bankOutFee)}`} dim />}
              {bIn > 0 && <CalcRow label={`− ${bankIn} ${(bIn * 100).toFixed(2).replace(".", ",")}%`} value={`−${fmtARS(calc.bankInFee)}`} dim />}
            </div>
          </div>
        )}
      </div>
    </Sheet>
  );
}
function BankSelect({ label, value, onChange }) {
  return (
    <div>
      <div style={{ fontSize: 10, color: C.muted, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.14em", fontWeight: 700 }}>{label}</div>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ width: "100%", padding: "12px 14px", borderRadius: 11, background: "rgba(255,255,255,0.03)", border: `1px solid ${C.border}`, color: C.text, fontSize: 14, fontFamily: "inherit" }}>
        {BANKS.map(b => <option key={b.id} value={b.id}>{b.label}</option>)}
      </select>
    </div>
  );
}
function CalcRow({ label, value, dim }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", fontSize: 12 }}>
      <span style={{ color: dim ? C.muted : C.soft }}>{label}</span>
      <span style={{ fontFamily: "'JetBrains Mono', monospace", color: dim ? C.muted : C.text, fontWeight: 600 }}>{value}</span>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET · SETTINGS / BACKUP
// ═══════════════════════════════════════════════════════════════════
function SettingsSheet({ ops, prefs, onClose, onImport, notify }) {
  const fileInput = useRef(null);

  const doExport = () => {
    exportBackup(ops, prefs);
    notify("✓ Backup descargado");
  };
  const doImport = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (!data.ops) throw new Error("Formato inválido");
        if (!confirm(`Restaurar ${data.ops.length} operaciones? Las actuales se reemplazarán.`)) return;
        onImport(data);
        onClose();
      } catch (err) { notify("Archivo inválido", { kind: "error" }); }
    };
    reader.readAsText(file);
  };

  return (
    <Sheet title="Más" subtitle="Backup, restaurar, info" onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <Row icon={Database} label="Backup completo" desc="Descargá JSON con todo (ops + ajustes)" onClick={doExport} />
        <Row icon={Upload} label="Restaurar backup" desc="Cargá un JSON exportado previamente" onClick={() => fileInput.current?.click()} />
        <input ref={fileInput} type="file" accept=".json" onChange={doImport} style={{ display: "none" }} />

        <div style={{ marginTop: 16, padding: "16px 18px", background: C.accentDim, border: `1px solid ${C.accent}33`, borderRadius: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
            <AtomLogo size={32} />
            <div>
              <div style={{ fontSize: 18, fontWeight: 800, letterSpacing: "-0.02em" }}>Plexo</div>
              <div style={{ fontSize: 10, color: C.muted, letterSpacing: "0.18em", textTransform: "uppercase", fontWeight: 600 }}>v1.0 · Operaciones</div>
            </div>
          </div>
          <div style={{ fontSize: 12, color: C.textDim, lineHeight: 1.5 }}>
            Operativa cripto/ARS sin fricción. Tus datos viven en tu dispositivo. Hacé backup frecuente.
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 14px", borderRadius: 11, background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}` }}>
          <span style={{ fontSize: 11, color: C.muted, letterSpacing: "0.1em", textTransform: "uppercase", fontWeight: 700 }}>Operaciones</span>
          <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, color: C.text, fontWeight: 700 }}>{ops.length}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 14px", borderRadius: 11, background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}` }}>
          <span style={{ fontSize: 11, color: C.muted, letterSpacing: "0.1em", textTransform: "uppercase", fontWeight: 700 }}>TTs cargadas</span>
          <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 14, color: C.text, fontWeight: 700 }}>{ops.reduce((s, o) => s + o.tts.length, 0)}</span>
        </div>
      </div>
    </Sheet>
  );
}
function Row({ icon: Icon, label, desc, onClick }) {
  return (
    <button onClick={onClick} style={{
      width: "100%", padding: "14px 16px", borderRadius: 12,
      background: "rgba(255,255,255,0.02)", border: `1px solid ${C.border}`,
      display: "flex", alignItems: "center", gap: 14, fontFamily: "inherit",
      transition: `all 0.2s ${EASE}`,
    }}>
      <div style={{ width: 38, height: 38, borderRadius: 11, background: C.accentDim, border: `1px solid ${C.accent}33`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <Icon size={16} color={C.accent} strokeWidth={2} />
      </div>
      <div style={{ flex: 1, textAlign: "left" }}>
        <div style={{ fontSize: 14, color: C.text, fontWeight: 600 }}>{label}</div>
        <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{desc}</div>
      </div>
      <ChevronRight size={16} color={C.muted} />
    </button>
  );
}
