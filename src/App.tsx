/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  Search, 
  Download, 
  Shuffle, 
  AlertCircle, 
  CheckCircle2, 
  FileSpreadsheet, 
  User, 
  ClipboardList,
  RefreshCw,
  ExternalLink
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel } from 'docx';

// --- Types ---
interface Student {
  id: string;
  name: string;
  gradeClass: string; // e.g., "1-3"
}

interface TopicState {
  [studentId: string]: string;
}

// --- Utils ---
const getCsvUrl = (url: string) => {
  if (!url) return '';
  try {
    // Convert regular spreadsheet URL to CSV export link
    // https://docs.google.com/spreadsheets/d/[ID]/edit#gid=0 -> https://docs.google.com/spreadsheets/d/[ID]/export?format=csv
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match) {
      return `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv`;
    }
    return url;
  } catch (e) {
    return url;
  }
};

const shuffleArray = <T,>(array: T[]): T[] => {
  const result = [...array];
  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
};

// --- Components ---

export default function App() {
  const [sheetUrl, setSheetUrl] = useState<string>(() => localStorage.getItem('ppt_manager_url') || '');
  const [allStudents, setAllStudents] = useState<Student[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  const [selectedClass, setSelectedClass] = useState<string>('');
  const [topics, setTopics] = useState<TopicState>({});
  const [displayOrder, setDisplayOrder] = useState<string[]>([]);
  const [isRandomized, setIsRandomized] = useState(false);

  // Grouped classes for dropdown
  const classes = useMemo(() => {
    const set = new Set<string>();
    allStudents.forEach(s => set.add(s.gradeClass));
    return Array.from(set).sort();
  }, [allStudents]);

  // Current filtered students
  const filteredStudents = useMemo(() => {
    return allStudents.filter(s => s.gradeClass === selectedClass);
  }, [allStudents, selectedClass]);

  // Sync display order when class changes or if not randomized
  useEffect(() => {
    if (!isRandomized) {
      setDisplayOrder(filteredStudents.map(s => s.id));
    }
  }, [filteredStudents, isRandomized]);

  // Duplicate Check
  const duplicateTopics = useMemo(() => {
    const counts: Record<string, number> = {};
    (Object.values(topics) as string[]).forEach(t => {
      const trimmed = t.trim();
      if (trimmed) {
        counts[trimmed] = (counts[trimmed] || 0) + 1;
      }
    });
    return new Set(Object.keys(counts).filter(k => counts[k] > 1));
  }, [topics]);

  // Fetch Logic
  const fetchSheetData = async (url: string) => {
    if (!url) return;
    setLoading(true);
    setError(null);
    localStorage.setItem('ppt_manager_url', url);

    try {
      const csvUrl = getCsvUrl(url);
      const response = await fetch(csvUrl);
      if (!response.ok) throw new Error('시트 데이터를 가져오는데 실패했습니다. 공유 설정을 확인하세요.');
      
      const csvText = await response.text();
      Papa.parse(csvText, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const data = results.data as any[];
          // Expected columns: 学年-班 (Grade-Class), 姓名 (Name) or similar
          // Mapping logic: try to find keys that contain "학년-반" and "이름"
          const students: Student[] = data.map((row, idx) => {
            const gradeKey = Object.keys(row).find(k => k === '학년' || k.includes('Grade') || k.includes('학년도'));
            const classOnlyKey = Object.keys(row).find(k => k === '반' || k === '학급' || k.includes('Class'));
            const combinedKey = Object.keys(row).find(k => k.includes('학년-반') || k.includes('학년반') || k.includes('Grade-Class'));
            
            let gradeClass = '기타';
            if (combinedKey && row[combinedKey]) {
              gradeClass = row[combinedKey].toString().trim();
            } else if (gradeKey && classOnlyKey && row[gradeKey] && row[classOnlyKey]) {
              gradeClass = `${row[gradeKey].toString().trim()}-${row[classOnlyKey].toString().trim()}`;
            } else {
              const fallbackKey = Object.keys(row).find(k => k.includes('반') || k.includes('학급') || k.includes('Grade') || k.includes('Class'));
              gradeClass = (row[fallbackKey || '학년-반'] || '기타').toString().trim();
            }
            
            const nameKey = Object.keys(row).find(k => k.includes('이름') || k.includes('성명') || k.includes('Name'));
            
            return {
              id: `student-${idx}`,
              name: (row[nameKey || '이름'] || row['성명'] || row['Name'] || '무명').toString().trim(),
              gradeClass: gradeClass
            };
          });
          
          if (students.length === 0) throw new Error('시트에서 학생 명단을 찾을 수 없습니다.');
          
          setAllStudents(students);
          setLoading(false);
          // Auto-select first class if none selected
          const firstClass = Array.from(new Set(students.map(s => s.gradeClass))).sort()[0];
          if (firstClass) setSelectedClass(firstClass);
        },
        error: (err: any) => {
          setError(`CSV 파싱 에러: ${err.message}`);
          setLoading(false);
        }
      });
    } catch (err: any) {
      setError(err.message);
      setLoading(false);
    }
  };

  const handleTopicChange = (id: string, value: string) => {
    setTopics(prev => ({ ...prev, [id]: value }));
  };

  const handleRandomize = () => {
    // Only randomize students who have topics entered? Or all?
    // Requirement says "주제 입력 학생들을 무작위 순서로 정렬"
    const studentsWithTopics = filteredStudents.filter((s: Student) => topics[s.id]?.trim());
    const studentsWithoutTopics = filteredStudents.filter((s: Student) => !topics[s.id]?.trim());
    
    if (studentsWithTopics.length === 0) {
      alert("발표 주제가 입력된 학생이 없습니다.");
      return;
    }

    const shuffled = shuffleArray<Student>(studentsWithTopics);
    setDisplayOrder([...shuffled.map((s: Student) => s.id), ...studentsWithoutTopics.map((s: Student) => s.id)]);
    setIsRandomized(true);
  };

  const handleResetOrder = () => {
    setDisplayOrder(filteredStudents.map(s => s.id));
    setIsRandomized(false);
  };

  const downloadDocx = async () => {
    const orderedStudents = displayOrder
      .map(id => filteredStudents.find(s => s.id === id))
      .filter((s): s is Student => !!s && (topics[s.id]?.trim() ? true : false));

    if (orderedStudents.length === 0) {
      alert("다운로드할 발표 결과가 없습니다.");
      return;
    }

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: `${selectedClass} 발표 순서 명단`,
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ text: "순번", alignment: AlignmentType.CENTER })] }),
                  new TableCell({ children: [new Paragraph({ text: "이름", alignment: AlignmentType.CENTER })] }),
                  new TableCell({ children: [new Paragraph({ text: "발표 주제", alignment: AlignmentType.CENTER })] }),
                ]
              }),
              ...orderedStudents.map((s, index) => new TableRow({
                children: [
                   new TableCell({ children: [new Paragraph({ text: (index + 1).toString(), alignment: AlignmentType.CENTER })] }),
                   new TableCell({ children: [new Paragraph({ text: s.name, alignment: AlignmentType.CENTER })] }),
                   new TableCell({ children: [new Paragraph({ text: topics[s.id] || "" })] }),
                ]
              }))
            ]
          })
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${selectedClass}_발표명단.docx`);
  };

  return (
    <div className="flex flex-col h-screen bg-[#0a0a0a] text-neutral-200 font-sans overflow-hidden">
      {/* Top Navigation Bar */}
      <header className="h-16 border-b border-neutral-800 flex items-center justify-between px-8 bg-[#0f0f0f] shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
            <Shuffle className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-semibold tracking-tight text-white">Presentation Master Pro</h1>
        </div>
        
        <div className="flex items-center gap-6">
          <div className="hidden md:flex flex-col items-end mr-2">
            <span className="text-[10px] uppercase tracking-widest text-neutral-500 font-bold">Students Count</span>
            <span className="text-sm text-neutral-300 italic font-mono">{filteredStudents.length}</span>
          </div>
          
          <div className="h-8 w-px bg-neutral-800 mx-2 hidden md:block"></div>
          
          <div className="flex items-center gap-3">
             <select 
              value={selectedClass}
              onChange={(e) => setSelectedClass(e.target.value)}
              className="bg-neutral-900 border border-neutral-700 text-neutral-200 rounded-md px-4 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer appearance-none min-w-[180px]"
            >
              <option value="" disabled>Select Class</option>
              {classes.map(c => (
                <option key={c} value={c}>{c}</option>
              ))}
              {classes.length === 0 && <option disabled>No data loaded</option>}
            </select>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 flex overflow-hidden">
        
        {/* Left Pane: Student List & Topic Input */}
        <section className="flex-1 border-r border-neutral-800 flex flex-col bg-[#0a0a0a]">
          <div className="p-4 border-b border-neutral-800 bg-[#0d0d0d] flex justify-between items-center shrink-0">
            <div className="flex items-center gap-3">
              <h2 className="text-xs font-bold uppercase tracking-widest text-neutral-500">Student Roster & Topics</h2>
              {isRandomized && (
                <span className="text-[10px] bg-blue-500/10 text-blue-400 px-2 py-0.5 rounded-full border border-blue-500/20">
                  RANDOMIZED
                </span>
              )}
            </div>
            
            {duplicateTopics.size > 0 && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex items-center gap-2 text-xs text-red-400 font-medium bg-red-950/30 px-3 py-1 rounded-full border border-red-900/50"
              >
                <span className="w-2 h-2 bg-red-500 rounded-full animate-pulse shadow-[0_0_8px_rgba(239,68,68,0.6)]"></span>
                Duplicate Topics Detected
              </motion.div>
            )}
          </div>
          
          <div className="flex-1 overflow-y-auto p-6 space-y-3 custom-scrollbar">
            {/* Table Header */}
            <div className="grid grid-cols-12 gap-4 px-4 py-2 text-[10px] uppercase font-bold text-neutral-600 tracking-wider sticky top-0 bg-[#0a0a0a] z-10 -mt-2 mb-2">
              <div className="col-span-1 text-center font-mono">No.</div>
              <div className="col-span-3">Name</div>
              <div className="col-span-8">Presentation Topic</div>
            </div>

            {/* List Items */}
            <AnimatePresence mode="popLayout" initial={false}>
              {displayOrder.map((id, index) => {
                const student = filteredStudents.find(s => s.id === id);
                if (!student) return null;
                const topic = topics[id] || "";
                const isDupe = topic.trim() && duplicateTopics.has(topic.trim());

                return (
                  <motion.div 
                    layout
                    key={id}
                    initial={{ opacity: 0, x: -20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    className={`grid grid-cols-12 gap-4 items-center p-3 rounded-lg border transition-all ${
                      isDupe 
                        ? 'bg-red-900/10 border-red-500/30 ring-1 ring-red-500/20' 
                        : topic.trim() 
                          ? 'bg-[#111111] border-neutral-800/80 hover:border-neutral-700' 
                          : 'bg-[#0c0c0c] border-neutral-900 opacity-60 hover:opacity-100 hover:border-neutral-800'
                    }`}
                  >
                    <div className={`col-span-1 text-center font-mono text-xs ${isDupe ? 'text-red-400' : 'text-neutral-500'}`}>
                      {String(index + 1).padStart(2, '0')}
                    </div>
                    <div className="col-span-3 font-medium text-white truncate px-1">
                      {student.name}
                    </div>
                    <div className="col-span-8 relative">
                      <input 
                        type="text" 
                        placeholder="Enter presentation topic..." 
                        value={topic}
                        onChange={(e) => handleTopicChange(id, e.target.value)}
                        className={`w-full bg-neutral-950/50 border rounded px-3 py-1.5 text-sm outline-none transition-all ${
                          isDupe 
                            ? 'border-red-500/50 text-red-200 placeholder:text-red-900/50' 
                            : 'border-neutral-800 focus:border-blue-500 text-neutral-100 placeholder:text-neutral-800'
                        }`}
                      />
                      {isDupe && (
                        <div className="absolute -right-1 -top-1">
                          <AlertCircle className="w-3 h-3 text-red-500 fill-red-500/10" />
                        </div>
                      )}
                    </div>
                  </motion.div>
                );
              })}
            </AnimatePresence>

            {filteredStudents.length === 0 && (
              <div className="flex flex-col items-center justify-center h-64 text-neutral-700 opacity-30 select-none">
                <FileSpreadsheet className="w-16 h-16 mb-4" />
                <p className="text-sm font-medium tracking-widest uppercase">Connect sheet to view roster</p>
              </div>
            )}
          </div>
        </section>

        {/* Right Pane: Controls & Preview */}
        <section className="w-[380px] flex flex-col bg-[#080808]">
          <div className="p-6 space-y-8 flex-1 overflow-y-auto">
            
            {/* Sheet Link Section */}
            <div className="space-y-3">
               <h3 className="text-xs font-bold uppercase tracking-widest text-neutral-500 flex items-center gap-2">
                <RefreshCw className="w-3 h-3" />
                Connection
              </h3>
              <div className="bg-[#111111] border border-neutral-800 p-4 rounded-xl space-y-4">
                <div className="relative">
                  <input 
                    type="text" 
                    placeholder="G-Sheet Public URL"
                    value={sheetUrl}
                    onChange={(e) => setSheetUrl(e.target.value)}
                    className="w-full bg-neutral-950 border border-neutral-800 rounded px-3 py-2 text-xs focus:border-blue-500 outline-none text-neutral-400 placeholder:text-neutral-700"
                  />
                </div>
                <button 
                  onClick={() => fetchSheetData(sheetUrl)}
                  disabled={loading || !sheetUrl}
                  className="w-full bg-neutral-800 hover:bg-neutral-700 text-white font-semibold py-2 px-4 rounded text-xs flex items-center justify-center gap-2 transition-all disabled:opacity-30"
                >
                  {loading ? <RefreshCw className="w-3 h-3 animate-spin" /> : <Search className="w-3 h-3" />}
                  Reload Data
                </button>
                {error && <p className="text-[10px] text-red-500 mt-1">{error}</p>}
              </div>
            </div>

            {/* Action Panel */}
            <div className="space-y-4">
              <h3 className="text-xs font-bold uppercase tracking-widest text-neutral-500">Action Panel</h3>
              <div className="bg-[#111111] border border-neutral-800 p-6 rounded-xl space-y-3 shadow-2xl">
                <button 
                  onClick={handleRandomize}
                  disabled={filteredStudents.length === 0}
                  className="w-full bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white font-bold py-3 px-4 rounded-lg flex items-center justify-center gap-2 transition-all transform active:scale-[0.98]"
                >
                  <Shuffle className="w-4 h-4" />
                  Shuffle Random Order
                </button>
                <button 
                  onClick={downloadDocx}
                  disabled={filteredStudents.length === 0}
                  className="w-full bg-neutral-800 hover:bg-neutral-700 disabled:opacity-50 text-neutral-200 font-bold py-3 px-4 rounded-lg flex items-center justify-center gap-2 transition-all"
                >
                  <Download className="w-4 h-4" />
                  Export to DOCX
                </button>
                {isRandomized && (
                  <button 
                    onClick={handleResetOrder}
                    className="w-full py-2 text-[10px] text-neutral-600 hover:text-neutral-400 uppercase tracking-widest font-bold transition-colors"
                  >
                    Reset to Roster Order
                  </button>
                )}
              </div>
            </div>

            {/* Result Preview */}
            <div className="space-y-4 pb-8">
              <h3 className="text-xs font-bold uppercase tracking-widest text-neutral-500 flex justify-between items-center">
                Presentation Queue
                <span className="text-[10px] lowercase italic font-normal text-neutral-700">showing next 4</span>
              </h3>
              
              <div className="space-y-3">
                {displayOrder.slice(0, 4).map((id, index) => {
                  const student = filteredStudents.find(s => s.id === id);
                  if (!student) return null;
                  const topic = topics[id] || "";
                  const rank = index === 0 ? '1st' : index === 1 ? '2nd' : index === 2 ? '3rd' : `${index + 1}th`;
                  
                  return (
                    <div 
                      key={`preview-${id}`} 
                      className={`flex items-center gap-4 bg-[#111111] p-3 rounded border-l-4 transition-all ${
                        index === 0 ? 'border-blue-500 scale-[1.02] shadow-lg bg-[#151515]' : 'border-neutral-800/50'
                      }`}
                    >
                      <span className={`text-lg font-bold w-10 text-center ${index === 0 ? 'text-blue-500' : 'text-neutral-700'}`}>
                        {rank}
                      </span>
                      <div className="flex-1 min-w-0">
                        <p className={`text-sm font-bold truncate ${index === 0 ? 'text-white' : 'text-neutral-400'}`}>
                          {student.name}
                        </p>
                        <p className="text-[11px] text-neutral-600 italic truncate italic">
                          {topic.trim() || 'Pending topic...'}
                        </p>
                      </div>
                    </div>
                  );
                })}
                
                {filteredStudents.length === 0 && (
                   <div className="bg-[#0c0c0c] border border-neutral-900 border-dashed p-8 rounded text-center opacity-20">
                     <ClipboardList className="w-8 h-8 mx-auto mb-2" />
                     <p className="text-[10px] uppercase font-bold tracking-tighter">Queue Empty</p>
                   </div>
                )}
              </div>
            </div>
          </div>

          {/* Footer Status Panel */}
          <div className="border-t border-neutral-800 p-4 bg-[#0d0d0d] flex items-center justify-between shrink-0">
            <span className="text-[10px] text-neutral-500 font-mono tracking-tight">
              Total: {filteredStudents.length} Students | Ready: {(Object.values(topics) as string[]).filter(t => t.trim()).length}
            </span>
            <div className="flex items-center gap-2">
              <div className={`w-1.5 h-1.5 rounded-full ${allStudents.length > 0 ? 'bg-green-500' : 'bg-neutral-800 animate-pulse'}`}></div>
              <span className="text-[10px] text-neutral-500 uppercase tracking-widest font-bold">
                {allStudents.length > 0 ? 'Sheets Active' : 'Disconnected'}
              </span>
            </div>
          </div>
        </section>
      </main>
    </div>
  );

}
