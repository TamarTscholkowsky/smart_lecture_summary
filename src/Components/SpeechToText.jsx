import React, { useEffect, useRef, useState } from "react";
import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import { Document, Packer, Paragraph, TextRun } from "docx";
import {
    Box,
    Button,
    Container,
    Typography,
    Stack,
    Paper,
    ButtonGroup,
    Tooltip
} from "@mui/material";
import PlayArrowIcon from '@mui/icons-material/PlayArrow';
import StopIcon from '@mui/icons-material/Stop';
import DeleteIcon from '@mui/icons-material/Delete';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import FormatBoldIcon from '@mui/icons-material/FormatBold';
import FormatUnderlinedIcon from '@mui/icons-material/FormatUnderlined';

const SpeechToText = () => {
    const [listening, setListening] = useState(false);
    const recognitionRef = useRef(null);

    const editorContentBeforeSession = useRef("");
    const finalizedDuringSession = useRef("");

    const editor = useEditor({
        extensions: [
            StarterKit,
            Underline,
        ],
        content: '',
        editorProps: {
            attributes: {
                style: 'direction: rtl; text-align: right; min-height: 250px; padding: 15px; outline: none; border: 1px solid #ccc; border-radius: 4px;',
            },
        },
    });

    useEffect(() => {
        return () => {
            if (recognitionRef.current) {
                recognitionRef.current.stop();
            }
        };
    }, []);

    const startListening = () => {
        if (!editor) {
            alert("העורך לא מוכן עדיין");
            return;
        }

        editorContentBeforeSession.current = editor.isEmpty ? "" : editor.getHTML();
        finalizedDuringSession.current = "";

        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        if (!SpeechRecognition) {
            alert("הדפדפן שלך לא תומך בזיהוי קולי.");
            return;
        }

        recognitionRef.current = new SpeechRecognition();
        recognitionRef.current.lang = "he-IL";
        recognitionRef.current.continuous = true;
        recognitionRef.current.interimResults = true;

        recognitionRef.current.onresult = (event) => {
            let interimTranscript = "";

            for (let i = event.resultIndex; i < event.results.length; ++i) {
                const transcript = event.results[i][0].transcript;
                if (event.results[i].isFinal) {
                    finalizedDuringSession.current += transcript + " ";
                } else {
                    interimTranscript = transcript;
                }
            }

            const fullContent = editorContentBeforeSession.current +
                finalizedDuringSession.current +
                interimTranscript;

            editor.commands.setContent(fullContent);
            editor.commands.focus('end');
        };

        recognitionRef.current.onend = () => setListening(false);
        recognitionRef.current.onerror = (event) => {
            console.error("Speech recognition error", event.error);
            setListening(false);
        };

        recognitionRef.current.start();
        setListening(true);
    };

    const stopListening = () => {
        if (recognitionRef.current) {
            recognitionRef.current.stop();
        }
        setListening(false);
    };

    const isParentBold = (node) => {
        let parent = node.parentElement;
        while (parent) {
            if (parent.tagName === 'STRONG' || parent.tagName === 'B') return true;
            parent = parent.parentElement;
        }
        return false;
    };

    const isParentUnderline = (node) => {
        let parent = node.parentElement;
        while (parent) {
            if (parent.tagName === 'U') return true;
            parent = parent.parentElement;
        }
        return false;
    };

    const parseStyledNode = (node, runs, style) => {
        const text = node.textContent;
        if (text) {
            runs.push(new TextRun({ text, ...style }));
        }
    };

    const parseNode = (node, runs) => {
        node.childNodes.forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                const text = child.textContent;
                if (text) {
                    runs.push(new TextRun({
                        text: text,
                        bold: isParentBold(child),
                        underline: isParentUnderline(child) ? {} : undefined,
                    }));
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                if (child.tagName === 'STRONG' || child.tagName === 'B') {
                    parseStyledNode(child, runs, { bold: true });
                } else if (child.tagName === 'U') {
                    parseStyledNode(child, runs, { underline: {} });
                } else {
                    parseNode(child, runs);
                }
            }
        });
    };

    const parseHTMLToDocx = (html) => {
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        const paragraphs = [];

        tempDiv.querySelectorAll('p').forEach(p => {
            const runs = [];
            parseNode(p, runs);
            paragraphs.push(new Paragraph({
                children: runs.length > 0 ? runs : [new TextRun(" ")],
                rightToLeft: true,
            }));
        });

        return paragraphs.length > 0 ? paragraphs : [new Paragraph(" ")];
    };

    const exportToWord = async () => {
        if (!editor) return;
        const htmlContent = editor.getHTML();

        if (!htmlContent || htmlContent === '<p></p>' || editor.isEmpty) {
            alert("אין טקסט לייצוא");
            return;
        }

        const paragraphs = parseHTMLToDocx(htmlContent);

        const doc = new Document({
            sections: [{
                properties: {
                    rightToLeft: true,
                },
                children: paragraphs,
            }],
        });

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `speech_to_text_${new Date().toISOString().slice(0, 10)}.docx`; // ✅ שם קובץ עם תאריך
        a.click();
        URL.revokeObjectURL(url);
    };

    return (
        <Container maxWidth="md" sx={{ mt: 8 }}>
            <Paper elevation={6} sx={{ p: 4, borderRadius: 3 }}>
                <Typography variant="h4" align="center" gutterBottom sx={{ fontWeight: 'bold', color: '#1976d2' }}>
                    המרה לטקסט עם עיצוב
                </Typography>

                <Stack direction="row" spacing={2} justifyContent="center" mb={3} flexWrap="wrap">
                    <Tooltip title={listening ? "המערכת מקשיבה כעת" : "לחץ להתחלת הקלטה"}>
                        <Button variant="contained" color={listening ? "secondary" : "success"} startIcon={<PlayArrowIcon />} onClick={startListening} disabled={listening}>
                            {listening ? "מקשיב..." : "התחל דיבור"}
                        </Button>
                    </Tooltip>
                    <Tooltip title="עצור הקלטה">
                        <Button variant="contained" color="error" startIcon={<StopIcon />} onClick={stopListening} disabled={!listening}>עצור</Button>
                    </Tooltip>
                    <Tooltip title="מחק את כל הטקסט">
                        <Button variant="contained" color="primary" startIcon={<DeleteIcon />} onClick={() => {
                            editor?.commands.clearContent();
                            editorContentBeforeSession.current = "";
                            finalizedDuringSession.current = "";
                        }}>נקה הכל</Button>
                    </Tooltip>
                    <Tooltip title="הורד כקובץ Word עם שמירת עיצוב">
                        <Button variant="contained" color="warning" startIcon={<FileDownloadIcon />} onClick={exportToWord}>הורד Word</Button>
                    </Tooltip>
                </Stack>

                <Box mb={1} textAlign="right">
                    <ButtonGroup size="small">
                        <Tooltip title="טקסט מודגש (Bold)">
                            <Button
                                onClick={() => editor?.chain().focus().toggleBold().run()}
                                variant={editor?.isActive('bold') ? 'contained' : 'outlined'}
                            >
                                <FormatBoldIcon />
                            </Button>
                        </Tooltip>

                        <Tooltip title="קו תחתון (Underline)">
                            <Button
                                onClick={() => editor?.chain().focus().toggleUnderline().run()}
                                variant={editor?.isActive('underline') ? 'contained' : 'outlined'}
                            >
                                <FormatUnderlinedIcon />
                            </Button>
                        </Tooltip>
                    </ButtonGroup>
                </Box>

                <Box sx={{
                    bgcolor: 'white',
                    borderRadius: 1,
                    "& .ProseMirror": {
                        minHeight: '250px',
                        direction: 'rtl',
                        textAlign: 'right'
                    }
                }}>
                    <EditorContent editor={editor} />
                </Box>
            </Paper>
        </Container>
    );
};

export default SpeechToText;