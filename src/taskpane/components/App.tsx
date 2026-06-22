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
