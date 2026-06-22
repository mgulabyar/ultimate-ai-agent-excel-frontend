// import React, { useState, useRef, useEffect } from "react";
// import {
//   Box,
//   Typography,
//   Button,
//   Paper,
//   CircularProgress,
//   Stack,
//   TextField,
//   IconButton,
//   Card,
//   Fade,
//   Divider,
//   Avatar,
//   Chip,
// } from "@mui/material";
// import {
//   SmartToy,
//   Send,
//   CheckCircle,
//   Terminal,
//   PlayCircleFilled,
//   Hub,
//   SettingsSuggest,
//   Computer,
// } from "@mui/icons-material";
// import axios from "axios";

// // TypeScript Interfaces for Precision
// interface AgentAction {
//   type: "WRITE" | "FORMAT" | "CLEAR" | "REPLACE";
//   range: string;
//   values?: any[][];
//   style?: {
//     backgroundColor?: string;
//     fontColor?: string;
//     bold?: boolean;
//     fontSize?: number;
//   };
// }

// const App: React.FC = () => {
//   const [loading, setLoading] = useState(false);
//   const [prompt, setPrompt] = useState("");
//   const [chat, setChat] = useState<{ role: "user" | "agent"; text: string }[]>([]);
//   const [executionLogs, setExecutionLogs] = useState<
//     { msg: string; status: "pending" | "done" | "error" }[]
//   >([]);
//   const chatEndRef = useRef<null | HTMLDivElement>(null);

//   // Auto-scroll chat window
//   useEffect(() => {
//     chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
//   }, [chat, executionLogs]);

//   // --- THE CORE EXECUTION ENGINE ---
//   const executeExcelMission = async (actions: AgentAction[]) => {
//     try {
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();

//         for (const action of actions) {
//           // Validation: Check if range is valid Excel notation
//           if (!action.range || action.range.includes(" ")) continue;

//           const excelRange = sheet.getRange(action.range);

//           if (action.type === "WRITE") {
//             setExecutionLogs((prev) => [
//               ...prev,
//               { msg: `Updating data at ${action.range}`, status: "pending" },
//             ]);
//             excelRange.values = action.values || [[]];
//           }

//           if (action.type === "FORMAT" && action.style) {
//             setExecutionLogs((prev) => [
//               ...prev,
//               { msg: `Applying styles to ${action.range}`, status: "pending" },
//             ]);
//             if (action.style.backgroundColor)
//               excelRange.format.fill.color = action.style.backgroundColor;
//             if (action.style.fontColor) excelRange.format.font.color = action.style.fontColor;
//             if (action.style.bold !== undefined) excelRange.format.font.bold = action.style.bold;
//             if (action.style.fontSize) excelRange.format.font.size = action.style.fontSize;
//           }

//           excelRange.format.autofitColumns();
//           // Force sync after each individual action for sequential reliability
//           await context.sync();
//           setExecutionLogs((prev) => {
//             const updated = [...prev];
//             updated[updated.length - 1].status = "done";
//             return updated;
//           });
//         }
//       });
//     } catch (error) {
//       console.error(error);
//       setExecutionLogs((prev) => [...prev, { msg: "Excel Execution Error", status: "error" }]);
//     }
//   };

// const handleSend = async () => {
//     if (!prompt.trim()) return;

//     const userText = prompt;
//     setChat(prev => [...prev, { role: 'user', text: userText }]);
//     setPrompt("");
//     setLoading(true);
//     setExecutionLogs([]);

//     try {
//       let sheetSnapshot: any[][] = [];

//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const usedRange = sheet.getUsedRange();
//         usedRange.load("values");

//         // Error handling for empty sheets
//         await context.sync().catch(() => {
//           console.log("Empty sheet detected.");
//           sheetSnapshot = [];
//         });

//         if (usedRange.values) {
//           sheetSnapshot = usedRange.values;
//         }

//         // Backend Call with validated data
//         const response = await axios.post("http://127.0.0.1:8007/api/agent/chat", {
//           prompt: userText,
//           snapshot: sheetSnapshot
//         });

//         const agentData = response.data.agent_data;
//         setChat(prev => [...prev, { role: 'agent', text: agentData.message }]);

//         if (agentData.actions && agentData.actions.length > 0) {
//           await executeExcelMission(agentData.actions);
//         }
//       });
//     } catch (error) {
//       console.error(error);
//       setChat(prev => [...prev, { role: 'agent', text: "Agent is ready. Please ensure your backend is running on Port 8007." }]);
//     }
//     setLoading(false);
//   };
//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#f8fafc", display: "flex", flexDirection: "column" }}>
//       {/* 1. CLEAN ENTERPRISE HEADER */}
//       <Paper
//         elevation={0}
//         sx={{
//           p: 2,
//           bgcolor: "#ffffff",
//           borderBottom: "1px solid #e2e8f0",
//           display: "flex",
//           alignItems: "center",
//           justifyContent: "space-between",
//         }}
//       >
//         <Stack direction="row" spacing={1} alignItems="center">
//           <Avatar sx={{ bgcolor: "#2563eb", width: 28, height: 28 }}>
//             <Hub sx={{ fontSize: 18 }} />
//           </Avatar>
//           <Typography
//             variant="subtitle2"
//             fontWeight={800}
//             color="#1e293b"
//             sx={{ letterSpacing: 0.5 }}
//           >
//             SUPREME AGENT <span style={{ color: "#2563eb" }}>PRO</span>
//           </Typography>
//         </Stack>
//         <Chip
//           label="v8.0"
//           size="small"
//           variant="outlined"
//           sx={{ height: 20, fontSize: "10px", fontWeight: 700 }}
//         />
//       </Paper>

//       {/* 2. MAIN CHAT & LOGS INTERFACE */}
//       <Box
//         sx={{
//           flexGrow: 1,
//           p: 2,
//           overflowY: "auto",
//           display: "flex",
//           flexDirection: "column",
//           gap: 2,
//         }}
//       >
//         {chat.length === 0 && (
//           <Box sx={{ textAlign: "center", mt: 5, opacity: 0.3 }}>
//             <Computer sx={{ fontSize: 40, mb: 1 }} />
//             <Typography variant="body2">System Online. Awaiting Mission.</Typography>
//           </Box>
//         )}

//         {chat.map((msg, i) => (
//           <Box
//             key={i}
//             sx={{ alignSelf: msg.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}
//           >
//             <Paper
//               sx={{
//                 p: 1.5,
//                 borderRadius: 2.5,
//                 bgcolor: msg.role === "user" ? "#2563eb" : "#ffffff",
//                 color: msg.role === "user" ? "#ffffff" : "#334155",
//                 boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
//                 border: msg.role === "agent" ? "1px solid #e2e8f0" : "none",
//               }}
//             >
//               <Typography sx={{ fontSize: "12.5px", lineHeight: 1.6 }}>{msg.text}</Typography>
//             </Paper>
//           </Box>
//         ))}

//         {/* 3. DYNAMIC EXECUTION TERMINAL */}
//         {(loading || executionLogs.length > 0) && (
//           <Card
//             elevation={0}
//             sx={{ bgcolor: "#1e293b", p: 1.5, borderRadius: 2, border: "1px solid #334155" }}
//           >
//             <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
//               <Terminal sx={{ fontSize: 14, color: "#38bdf8" }} />
//               <Typography
//                 variant="caption"
//                 sx={{ color: "#38bdf8", fontWeight: 800, letterSpacing: 1 }}
//               >
//                 LIVE EXECUTION ENGINE
//               </Typography>
//             </Stack>
//             {executionLogs.map((log, i) => (
//               <Stack key={i} direction="row" spacing={1} alignItems="center" sx={{ mb: 0.5 }}>
//                 {log.status === "done" ? (
//                   <CheckCircle sx={{ fontSize: 10, color: "#10b981" }} />
//                 ) : (
//                   <CircularProgress size={8} sx={{ color: "#38bdf8" }} />
//                 )}
//                 <Typography
//                   sx={{
//                     fontFamily: "monospace",
//                     fontSize: "10.5px",
//                     color: log.status === "done" ? "#10b981" : "#cbd5e1",
//                   }}
//                 >
//                   {`> ${log.msg}`}
//                 </Typography>
//               </Stack>
//             ))}
//           </Card>
//         )}
//         <div ref={chatEndRef} />
//       </Box>

//       {/* 4. COMMAND INPUT PANEL */}
//       <Box sx={{ p: 2, bgcolor: "#ffffff", borderTop: "1px solid #e2e8f0" }}>
//         <TextField
//           fullWidth
//           size="small"
//           placeholder="Enter Excel instruction (Urdu/English)..."
//           value={prompt}
//           onChange={(e) => setPrompt(e.target.value)}
//           onKeyPress={(e) => e.key === "Enter" && handleSend()}
//           InputProps={{
//             endAdornment: (
//               <IconButton
//                 onClick={handleSend}
//                 disabled={loading}
//                 color="primary"
//                 sx={{ bgcolor: loading ? "transparent" : "#f1f5f9" }}
//               >
//                 {loading ? <CircularProgress size={20} /> : <Send sx={{ fontSize: 18 }} />}
//               </IconButton>
//             ),
//             sx: { borderRadius: 3, fontSize: "13px", bgcolor: "#f8fafc" },
//           }}
//         />
//         <Typography
//           variant="caption"
//           sx={{ display: "block", textAlign: "center", mt: 1, color: "#94a3b8", fontSize: "9px" }}
//         >
//           OMNI-CORE ACTIVE • PORT 8007 • GPT-4 TURBO
//         </Typography>
//       </Box>
//     </Box>
//   );
// };

// export default App;

// import React, { useState, useRef, useEffect } from "react";
// import {
//   Box,
//   Typography,
//   Button,
//   Paper,
//   CircularProgress,
//   Stack,
//   TextField,
//   IconButton,
//   Card,
//   Fade,
//   LinearProgress,
// } from "@mui/material";
// import {
//   SmartToy,
//   Send,
//   CheckCircle,
//   Terminal,
//   Settings,
//   Memory,
//   FactCheck,
//   DoneAll,
// } from "@mui/icons-material";
// import axios from "axios";

// const App: React.FC = () => {
//   const [loading, setLoading] = useState(false);
//   const [prompt, setPrompt] = useState("");
//   const [chat, setChat] = useState<{ role: string; text: string }[]>([]);
//   const [currentTasks, setCurrentTasks] = useState<{ msg: string; done: boolean }[]>([]);
//   const chatEndRef = useRef<null | HTMLDivElement>(null);

//   useEffect(() => {
//     chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
//   }, [chat, currentTasks]);

//   const executeActions = async (actions: any[]) => {
//     try {
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();

//         for (let i = 0; i < actions.length; i++) {
//           const action = actions[i];
//           setCurrentTasks((prev) => [
//             ...prev,
//             { msg: `Action ${i + 1}: ${action.type} on ${action.range}`, done: false },
//           ]);

//           const range = sheet.getRange(action.range);

//           if (action.type === "WRITE" || action.type === "FORMULA") {
//             range.values = action.values;
//           }

//           if (action.style) {
//             const fmt = range.format;
//             if (action.style.backgroundColor) fmt.fill.color = action.style.backgroundColor;
//             if (action.style.fontColor) fmt.font.color = action.style.fontColor;
//             if (action.style.bold) fmt.font.bold = true;

//             // ELITE FIX: Automatically apply Borders for tables
//             if (action.style.applyBorders || action.range.includes(":")) {
//               const borders = fmt.borders;
//               borders.load("items");
//               borders.getItem("EdgeTop").style = "Continuous";
//               borders.getItem("EdgeBottom").style = "Continuous";
//               borders.getItem("EdgeLeft").style = "Continuous";
//               borders.getItem("EdgeRight").style = "Continuous";
//               borders.getItem("InsideVertical").style = "Continuous";
//               borders.getItem("InsideHorizontal").style = "Continuous";
//             }
//           }

//           range.format.autofitColumns();
//           await context.sync();

//           setCurrentTasks((prev) => {
//             const updated = [...prev];
//             updated[updated.length - 1].done = true;
//             return updated;
//           });
//         }
//       });
//     } catch (e) {
//       console.error(e);
//     }
//   };

//   const handleSend = async () => {
//     if (!prompt.trim()) return;
//     const userText = prompt;
//     setChat((prev) => [...prev, { role: "user", text: userText }]);
//     setPrompt("");
//     setLoading(true);
//     setCurrentTasks([]);

//     try {
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const usedRange = sheet.getUsedRange();
//         usedRange.load("values");
//         await context.sync().catch(() => {});

//         const res = await axios.post("http://127.0.0.1:8007/api/agent/chat", {
//           prompt: userText,
//           snapshot: usedRange.values || [],
//         });

//         const { agent_data } = res.data;
//         setChat((prev) => [...prev, { role: "assistant", text: agent_data.message }]);
//         if (agent_data.actions) await executeActions(agent_data.actions);
//       });
//     } catch (e) {
//       setChat((prev) => [
//         ...prev,
//         { role: "assistant", text: "Brain sync interrupted. Check connection." },
//       ]);
//     }
//     setLoading(false);
//   };

//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#f8f9fa", display: "flex", flexDirection: "column" }}>
//       {/* 1. PROFESSIONAL HEADER */}
//       <Paper
//         elevation={0}
//         sx={{
//           p: 2,
//           bgcolor: "#fff",
//           borderBottom: "1px solid #dee2e6",
//           display: "flex",
//           alignItems: "center",
//           gap: 1.5,
//         }}
//       >
//         <Memory sx={{ color: "#2b579a" }} />
//         <Typography variant="subtitle2" fontWeight={800} color="#343a40">
//           OMNI-AGENT <span style={{ color: "#2b579a" }}>ELITE</span>
//         </Typography>
//       </Paper>

//       {/* 2. CHAT SCROLLER */}
//       <Box
//         sx={{
//           flexGrow: 1,
//           p: 2,
//           overflowY: "auto",
//           display: "flex",
//           flexDirection: "column",
//           gap: 2,
//         }}
//       >
//         {chat.map((m, i) => (
//           <Box
//             key={i}
//             sx={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}
//           >
//             <Paper
//               sx={{
//                 p: 1.5,
//                 borderRadius: 2.5,
//                 bgcolor: m.role === "user" ? "#2b579a" : "#fff",
//                 color: m.role === "user" ? "#fff" : "#495057",
//                 boxShadow: "0 2px 4px rgba(0,0,0,0.05)",
//                 border: m.role === "assistant" ? "1px solid #e9ecef" : "none",
//               }}
//             >
//               <Typography sx={{ fontSize: "13px", lineHeight: 1.6 }}>{m.text}</Typography>
//             </Paper>
//           </Box>
//         ))}

//         {/* 3. PROFESSIONAL EXECUTION LOGS */}
//         {(loading || currentTasks.length > 0) && (
//           <Card sx={{ bgcolor: "#212529", p: 2, borderRadius: 2, border: "1px solid #343a40" }}>
//             <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1.5 }}>
//               <FactCheck sx={{ fontSize: 16, color: "#0dcaf0" }} />
//               <Typography
//                 variant="caption"
//                 sx={{ color: "#0dcaf0", fontWeight: 800, letterSpacing: 0.5 }}
//               >
//                 TASK EXECUTION PIPELINE
//               </Typography>
//             </Stack>
//             <Stack spacing={1}>
//               {currentTasks.map((task, i) => (
//                 <Stack key={i} direction="row" spacing={1.5} alignItems="center">
//                   {task.done ? (
//                     <DoneAll sx={{ fontSize: 14, color: "#198754" }} />
//                   ) : (
//                     <CircularProgress size={10} sx={{ color: "#0dcaf0" }} />
//                   )}
//                   <Typography
//                     sx={{
//                       color: task.done ? "#adb5bd" : "#fff",
//                       fontFamily: "monospace",
//                       fontSize: "11px",
//                     }}
//                   >
//                     {task.msg}
//                   </Typography>
//                 </Stack>
//               ))}
//             </Stack>
//             {loading && (
//               <LinearProgress
//                 sx={{
//                   mt: 2,
//                   height: 2,
//                   bgcolor: "#343a40",
//                   "& .MuiLinearProgress-bar": { bgcolor: "#0dcaf0" },
//                 }}
//               />
//             )}
//           </Card>
//         )}
//         <div ref={chatEndRef} />
//       </Box>

//       {/* 4. COMPACT INPUT */}
//       <Box sx={{ p: 2, bgcolor: "#fff", borderTop: "1px solid #dee2e6" }}>
//         <TextField
//           fullWidth
//           size="medium"
//           placeholder="Type command..."
//           value={prompt}
//           onChange={(e) => setPrompt(e.target.value)}
//           onKeyPress={(e) => e.key === "Enter" && handleSend()}
//           InputProps={{
//             endAdornment: (
//               <IconButton onClick={handleSend} disabled={loading} color="primary">
//                 {loading ? <CircularProgress size={22} /> : <Send sx={{ fontSize: 18 }} />}
//               </IconButton>
//             ),
//             sx: { borderRadius: 2, fontSize: "13px", bgcolor: "#f1f3f5" },
//           }}
//         />
//         <Typography
//           variant="caption"
//           sx={{ display: "block", mt: 1, textAlign: "center", color: "#adb5bd", fontSize: "9px" }}
//         >
//           V 5.0 • CORE 8007 • BORDER LOGIC ACTIVE
//         </Typography>
//       </Box>
//     </Box>
//   );
// };

// export default App;

// import React, { useState, useRef, useEffect } from "react";
// import {
//   Box,
//   Typography,
//   Paper,
//   CircularProgress,
//   Stack,
//   TextField,
//   IconButton,
//   Avatar,
// } from "@mui/material";
// import { Send, FlashOn, Psychology } from "@mui/icons-material";
// import axios from "axios";

// const App: React.FC = () => {
//   const [loading, setLoading] = useState(false);
//   const [prompt, setPrompt] = useState("");
//   const [chat, setChat] = useState<{ role: string; text: string }[]>([]);
//   const chatEndRef = useRef<null | HTMLDivElement>(null);

//   useEffect(() => {
//     chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
//   }, [chat]);

//   // CORE EXECUTION ENGINE (Fixed TypeScript Overload Error)
//   const executeActions = async (actions: any[]) => {
//     try {
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();

//         for (const action of actions) {
//           const range = sheet.getRange(action.range);

//           if (action.type === "WRITE" || action.type === "FORMULA" || action.type === "REPLACE") {
//             range.values = action.values;
//           }

//           if (action.style) {
//             const fmt = range.format;
//             if (action.style.backgroundColor) fmt.fill.color = action.style.backgroundColor;
//             if (action.style.fontColor) fmt.font.color = action.style.fontColor;
//             if (action.style.bold !== undefined) fmt.font.bold = action.style.bold;

//             // --- FIXED BORDER LOGIC ---
//             if (action.style.applyBorders) {
//               // Defining explicit literal types for TypeScript
//               const borders: (
//                 | "EdgeTop"
//                 | "EdgeBottom"
//                 | "EdgeLeft"
//                 | "EdgeRight"
//                 | "InsideVertical"
//                 | "InsideHorizontal"
//               )[] = [
//                 "EdgeTop",
//                 "EdgeBottom",
//                 "EdgeLeft",
//                 "EdgeRight",
//                 "InsideVertical",
//                 "InsideHorizontal",
//               ];

//               borders.forEach((b) => {
//                 fmt.borders.getItem(b).style = "Continuous";
//                 fmt.borders.getItem(b).color = "#D1D5DB"; // Subtle gray borders
//               });
//             }

//             if (action.style.removeBorders) {
//               const borders: (
//                 | "EdgeTop"
//                 | "EdgeBottom"
//                 | "EdgeLeft"
//                 | "EdgeRight"
//                 | "InsideVertical"
//                 | "InsideHorizontal"
//               )[] = [
//                 "EdgeTop",
//                 "EdgeBottom",
//                 "EdgeLeft",
//                 "EdgeRight",
//                 "InsideVertical",
//                 "InsideHorizontal",
//               ];

//               borders.forEach((b) => {
//                 fmt.borders.getItem(b).style = "None";
//               });
//             }
//           }
//           range.format.autofitColumns();
//           await context.sync();
//         }
//       });
//     } catch (e) {
//       console.error("Execution Error:", e);
//     }
//   };

//   const handleSend = async () => {
//     if (!prompt.trim()) return;
//     const userText = prompt;
//     setChat((prev) => [...prev, { role: "user", text: userText }]);
//     setPrompt("");
//     setLoading(true);

//     try {
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const usedRange = sheet.getUsedRange();
//         usedRange.load("values");
//         await context.sync().catch(() => {});

//         const res = await axios.post("http://127.0.0.1:8007/api/agent/chat", {
//           prompt: userText,
//           snapshot: usedRange.values || [],
//         });

//         const { agent_data } = res.data;
//         setChat((prev) => [...prev, { role: "assistant", text: agent_data.message }]);
//         if (agent_data.actions) await executeActions(agent_data.actions);
//       });
//     } catch (e) {
//       setChat((prev) => [
//         ...prev,
//         { role: "assistant", text: "Connection error. Is Node 8007 running?" },
//       ]);
//     }
//     setLoading(false);
//   };

//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#f8f9fa", display: "flex", flexDirection: "column" }}>
//       {/* 1. CLEAN HEADER */}
//       <Paper
//         elevation={0}
//         sx={{
//           p: 1.5,
//           bgcolor: "#fff",
//           borderBottom: "1px solid #e9ecef",
//           display: "flex",
//           alignItems: "center",
//           justifyContent: "space-between",
//         }}
//       >
//         <Stack direction="row" spacing={1} alignItems="center">
//           <Avatar sx={{ bgcolor: "#2b579a", width: 30, height: 30 }}>
//             <Psychology sx={{ fontSize: 20 }} />
//           </Avatar>
//           <Typography variant="subtitle2" fontWeight={800} color="#343a40">
//             OMNI-COMMANDER AI
//           </Typography>
//         </Stack>
//         <FlashOn sx={{ fontSize: 18, color: "#2b579a", opacity: 0.6 }} />
//       </Paper>

//       {/* 2. CHAT AREA */}
//       <Box
//         sx={{
//           flexGrow: 1,
//           p: 2,
//           overflowY: "auto",
//           display: "flex",
//           flexDirection: "column",
//           gap: 2,
//         }}
//       >
//         {chat.map((m, i) => (
//           <Box
//             key={i}
//             sx={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}
//           >
//             <Paper
//               sx={{
//                 p: 1.5,
//                 borderRadius: 2.5,
//                 bgcolor: m.role === "user" ? "#2b579a" : "#fff",
//                 color: m.role === "user" ? "#fff" : "#495057",
//                 boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
//                 border: m.role === "assistant" ? "1px solid #e9ecef" : "none",
//               }}
//             >
//               <Typography sx={{ fontSize: "12.5px", lineHeight: 1.6 }}>{m.text}</Typography>
//             </Paper>
//           </Box>
//         ))}
//         <div ref={chatEndRef} />
//       </Box>



import React, { useState, useRef, useEffect } from "react";
import {
  Box,
  Typography,
  Paper,
  CircularProgress,
  Stack,
  TextField,
  IconButton,
  Avatar,
  Fade,
} from "@mui/material";
import { Send, FlashOn, AutoMode, Terminal, AutoFixHigh, Psychology } from "@mui/icons-material";
import axios from "axios";

const App: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [prompt, setPrompt] = useState("");
  const [chat, setChat] = useState<{ role: string; text: string }[]>([]);
  const chatEndRef = useRef<null | HTMLDivElement>(null);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chat]);

 const executeActions = async (actions: any[]) => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        for (const action of actions) {
          console.log("Executing Action:", action); // Debugging log
          const range = sheet.getRange(action.range);
          
          if (action.type === "WRITE" || action.type === "FORMULA" || action.type === "REPLACE") {
            // Excel expects 2D array [[val1, val2]]
            range.values = action.values;
          }

          if (action.style) {
            const fmt = range.format;
            if (action.style.backgroundColor) fmt.fill.color = action.style.backgroundColor;
            if (action.style.fontColor) fmt.font.color = action.style.fontColor;
            if (action.style.bold !== undefined) fmt.font.bold = action.style.bold;
            
            if (action.style.applyBorders) {
              const borders: ("EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal")[] = 
                ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
              borders.forEach(b => {
                fmt.borders.getItem(b).style = "Continuous";
                fmt.borders.getItem(b).color = "#D1D5DB";
              });
            }
          }
          range.format.autofitColumns();
          await context.sync(); 
        }
      });
    } catch (e: any) {
      console.error("Excel Logic Error:", e);
      alert("Excel Error: " + e.message);
    }
  };

  const handleSend = async () => {
    if (!prompt.trim()) return;
    const userText = prompt;
    setChat(prev => [...prev, { role: 'user', text: userText }]);
    setPrompt("");
    setLoading(true);

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("values");
        await context.sync().catch(() => {});

        const res = await axios.post("http://127.0.0.1:8007/api/agent/chat", {
          prompt: userText,
          snapshot: usedRange.values || []
        });

        const { agent_data } = res.data;
        console.log("Raw AI Data:", agent_data); // Yeh check karein console mein!

        setChat(prev => [...prev, { role: 'assistant', text: agent_data.message }]);

        if (agent_data.actions && agent_data.actions.length > 0) {
          await executeActions(agent_data.actions);
        } else {
          console.warn("AI logic gave 0 actions.");
          alert("AI ne message to diya par koi action nahi bheja.");
        }
      });
    } catch (e) {
      setChat(prev => [...prev, { role: 'assistant', text: "Connection error or Excel sync failed." }]);
    }
    setLoading(false);
  };

  return (
    <Box sx={{ height: "100vh", bgcolor: "#f8f9fa", display: "flex", flexDirection: "column" }}>
      {/* 1. UPDATED HEADING: NEXUS AGENT ELITE */}
      <Paper
        elevation={0}
        sx={{
          p: 1.5,
          bgcolor: "#fff",
          borderBottom: "1px solid #e9ecef",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}
      >
        <Stack direction="row" spacing={1} alignItems="center">
          <Avatar sx={{ bgcolor: "#1e293b", width: 30, height: 30 }}>
            <AutoMode sx={{ fontSize: 20, color: "#38bdf8" }} />
          </Avatar>
          <Typography
            variant="subtitle2"
            fontWeight={900}
            color="#1e293b"
            sx={{ letterSpacing: 0.5 }}
          >
            NEXUS <span style={{ color: "#2563eb" }}>AGENT</span> ELITE
          </Typography>
        </Stack>
        <FlashOn sx={{ fontSize: 18, color: "#2563eb", opacity: 0.8 }} />
      </Paper>

      {/* 2. CHAT AREA WITH WELCOME SCREEN */}
      <Box
        sx={{
          flexGrow: 1,
          p: 2,
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: 2,
        }}
      >
        {chat.length === 0 && !loading && (
          <Fade in={true}>
            <Box sx={{ mt: 8, textAlign: "center", px: 2 }}>
              <Psychology sx={{ fontSize: 50, color: "#cbd5e1", mb: 2 }} />
              <Typography variant="h6" fontWeight={700} color="#334155">
                Nexus Intelligence
              </Typography>
              <Typography variant="body2" color="#64748b" sx={{ mt: 1, lineHeight: 1.6 }}>
                Command your Excel data with Natural Language.
                
              </Typography>
            </Box>
          </Fade>
        )}

        {chat.map((m, i) => (
          <Box
            key={i}
            sx={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}
          >
            <Paper
              sx={{
                p: 1.5,
                borderRadius: 2.5,
                bgcolor: m.role === "user" ? "#2563eb" : "#fff",
                color: m.role === "user" ? "#fff" : "#334155",
                boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
                border: m.role === "assistant" ? "1px solid #e9ecef" : "none",
              }}
            >
              <Typography sx={{ fontSize: "12.5px", lineHeight: 1.6 }}>{m.text}</Typography>
            </Paper>
          </Box>
        ))}
        <div ref={chatEndRef} />
      </Box>

      {/* 3. INPUT PANEL */}
      <Box sx={{ p: 2, bgcolor: "#fff", borderTop: "1px solid #e9ecef" }}>
        <TextField
          fullWidth
          size="medium"
          placeholder="Instruct the agent..."
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          onKeyPress={(e) => e.key === "Enter" && handleSend()}
          InputProps={{
            endAdornment: (
              <IconButton onClick={handleSend} disabled={loading} color="primary">
                {loading ? <CircularProgress size={22} /> : <Send sx={{ fontSize: 18 }} />}
              </IconButton>
            ),
            sx: { borderRadius: 2, fontSize: "13px", bgcolor: "#f1f3f5" },
          }}
        />
        <Typography
          variant="caption"
          sx={{ display: "block", mt: 1, textAlign: "center", color: "#adb5bd", fontSize: "9px" }}
        >
          POWERED BY OMNI-BRAIN 8007 • BORDER LOGIC ACTIVE
        </Typography>
      </Box>
    </Box>
  );
};

export default App;
