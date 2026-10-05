import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import OpenAI from 'openai';
import mammoth from 'mammoth';
import { PDFParse } from 'pdf-parse';
import { Document, Packer, Paragraph, HeadingLevel, TextRun, Table, TableRow, TableCell, WidthType } from 'docx';
import pptxgen from 'pptxgenjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const DATA = path.join(ROOT, 'data');
const UPLOADS = path.join(ROOT, 'uploads');
const EXPORTS = path.join(ROOT, 'exports');
for (const d of [DATA, UPLOADS, EXPORTS]) fs.mkdirSync(d, { recursive: true });

const lessons = JSON.parse(fs.readFileSync(path.join(DATA, 'lessons.json'), 'utf8'));
const projectsFile = path.join(DATA, 'projects.json');
const sourcesFile = path.join(DATA, 'sources.json');
if (!fs.existsSync(projectsFile)) fs.writeFileSync(projectsFile, '[]');
if (!fs.existsSync(sourcesFile)) fs.writeFileSync(sourcesFile, '[]');
const readJSON = file => JSON.parse(fs.readFileSync(file, 'utf8'));
const writeJSON = (file, data) => fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
const readProjects = () => readJSON(projectsFile);
const readSources = () => readJSON(sourcesFile);
const id = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const app = express();
app.use(express.json({ limit: '8mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(ROOT, 'public')));

const upload = multer({
  dest: UPLOADS,
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (_, file, cb) => cb(null, /\.(pdf|docx|txt|md)$/i.test(file.originalname))
});

const client = process.env.OPENAI_API_KEY ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY }) : null;
const model = process.env.OPENAI_MODEL || 'gpt-6-luna';
const clean = s => String(s ?? '').replace(/\u0000/g, '').trim();
const clamp = (s, n = 50000) => clean(s).slice(0, n);
const ensureAI = () => { if (!client) throw new Error('Chưa cấu hình OPENAI_API_KEY trong file .env'); };

async function extractFile(file) {
  const ext = path.extname(file.originalname).toLowerCase();
  const buf = fs.readFileSync(file.path);
  try {
    if (ext === '.pdf') {
      const parser = new PDFParse({ data: buf });
      try {
        const result = await parser.getText();
        return clamp(result.text, 100000);
      } finally {
        await parser.destroy();
      }
    }
    if (ext === '.docx') return clamp((await mammoth.extractRawText({ buffer: buf })).value, 100000);
    return clamp(buf.toString('utf8'), 100000);
  } finally { fs.rmSync(file.path, { force: true }); }
}

const planSchema = {
  type: 'object', additionalProperties: false,
  properties: {
    title:{type:'string'}, grade:{type:'string'}, subject:{type:'string'}, duration:{type:'string'},
    objectives:{type:'array',items:{type:'string'}}, competencies:{type:'array',items:{type:'string'}}, qualities:{type:'array',items:{type:'string'}},
    preparation:{type:'object',additionalProperties:false,properties:{teacher:{type:'array',items:{type:'string'}},students:{type:'array',items:{type:'string'}}},required:['teacher','students']},
    activities:{type:'array',items:{type:'object',additionalProperties:false,properties:{phase:{type:'string'},time:{type:'string'},teacher:{type:'string'},students:{type:'string'},product:{type:'string'},assessment:{type:'string'}},required:['phase','time','teacher','students','product','assessment']}},
    knowledgeSummary:{type:'array',items:{type:'string'}},
    differentiation:{type:'object',additionalProperties:false,properties:{support:{type:'array',items:{type:'string'}},extension:{type:'array',items:{type:'string'}}},required:['support','extension']},
    homework:{type:'array',items:{type:'string'}},notes:{type:'array',items:{type:'string'}}
  }, required:['title','grade','subject','duration','objectives','competencies','qualities','preparation','activities','knowledgeSummary','differentiation','homework','notes']
};
const slideSchema = {
  type:'object',additionalProperties:false,
  properties:{title:{type:'string'},subtitle:{type:'string'},slides:{type:'array',items:{type:'object',additionalProperties:false,properties:{title:{type:'string'},bullets:{type:'array',items:{type:'string'}},speakerNotes:{type:'string'},activity:{type:'string'}},required:['title','bullets','speakerNotes','activity']}},teacherNotes:{type:'array',items:{type:'string'}}},
  required:['title','subtitle','slides','teacherNotes']
};

async function askAI(instructions, input, schema, name) {
  ensureAI();
  const r = await client.responses.create({
    model, store:false, instructions, input,
    text:{ format:{ type:'json_schema', name, strict:true, schema } }
  });
  if (!r.output_text) throw new Error('AI không trả về nội dung.');
  return JSON.parse(r.output_text);
}

function sourceContext(lessonId, extra='') {
  const lesson = lessons.find(x => x.id === lessonId);
  const saved = readSources().filter(s => !s.lessonId || s.lessonId === lessonId).slice(0, 8);
  const savedText = saved.map(s => `\n--- NGUỒN: ${s.name} | ${s.lessonId || 'chung'} ---\n${s.text}`).join('\n');
  return { lesson, saved, text: clamp(`${savedText}\n--- NGUỒN TỨC THỜI ---\n${extra}`, 90000) };
}

const commonRules = `Bạn là chuyên gia thiết kế dạy học Toán THCS Việt Nam. Sản phẩm dành cho giáo viên Toán 9, bộ Kết nối tri thức với cuộc sống. Chỉ khẳng định điều có trong metadata hoặc nguồn giáo viên cung cấp. Không bịa số trang, mục SGK, ví dụ hay định lý nếu nguồn không có. Nếu nguồn thiếu, dùng mô tả an toàn và ghi rõ giáo viên cần kiểm tra/bổ sung. Nội dung phải phù hợp học sinh lớp 9, tiếng Việt tự nhiên, khả thi trong lớp học.`;
const planInstruction = `${commonRules}\nTạo KẾ HOẠCH BÀI DẠY có mục tiêu, năng lực, phẩm chất, chuẩn bị, tiến trình, sản phẩm, đánh giá, phân hóa và dặn dò. Phân bổ thời gian phải hợp lý và tổng xấp xỉ thời lượng yêu cầu. Hoạt động phải có việc GV và HS cụ thể.`;
const slideInstruction = `${commonRules}\nTạo bộ slide để dạy trực tiếp. Slide ngắn, trực quan, không nhồi chữ. Nên có khởi động, hình thành kiến thức, ví dụ/luyện tập, hoạt động tương tác, vận dụng và chốt bài khi phù hợp. Speaker notes dành cho giáo viên.`;

app.get('/api/health', (_,res)=>res.json({ok:true,aiConfigured:!!client,model,version:'2.0'}));
app.get('/api/lessons', (_,res)=>res.json({lessons}));
app.get('/api/lessons/:id',(req,res)=>{const lesson=lessons.find(x=>x.id===req.params.id); if(!lesson)return res.status(404).json({error:'Không tìm thấy bài.'}); res.json({lesson});});

app.get('/api/sources',(req,res)=>{let items=readSources(); if(req.query.lessonId) items=items.filter(x=>!x.lessonId||x.lessonId===req.query.lessonId); res.json({sources:items.map(({text,...meta})=>({...meta,chars:text.length}))});});
app.post('/api/sources',upload.single('file'),async(req,res)=>{
  try{
    if(!req.file)return res.status(400).json({error:'Chưa chọn file.'});
    const text=await extractFile(req.file); const item={id:id(),name:req.file.originalname,lessonId:req.body.lessonId||'',createdAt:new Date().toISOString(),text};
    const all=readSources(); all.unshift(item); writeJSON(sourcesFile,all.slice(0,100)); res.json({source:{...item,text:undefined,chars:text.length}});
  }catch(e){res.status(400).json({error:e.message});}
});
app.delete('/api/sources/:id',(req,res)=>{writeJSON(sourcesFile,readSources().filter(x=>x.id!==req.params.id));res.json({ok:true});});

app.post('/api/generate-plan',async(req,res)=>{try{
  const {lessonId,duration=45,method='5E',request='',source=''}=req.body; const ctx=sourceContext(lessonId,source); if(!ctx.lesson)return res.status(400).json({error:'Bài học không hợp lệ.'});
  const input=`DỮ LIỆU BÀI HỌC:\n${JSON.stringify(ctx.lesson,null,2)}\n\nTHỜI LƯỢNG: ${duration} phút\nMÔ HÌNH: ${method}\nYÊU CẦU: ${request||'Không có'}\n\nKHO NGUỒN ĐÃ NẠP:\n${ctx.text||'(Không có nguồn chi tiết.)'}`;
  const data=await askAI(planInstruction,input,planSchema,'lesson_plan'); const project={id:id(),type:'plan',name:data.title||ctx.lesson.title,lessonId,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),data}; const all=readProjects();all.unshift(project);writeJSON(projectsFile,all.slice(0,200));res.json({project});
}catch(e){res.status(500).json({error:e.message});}});

app.post('/api/generate-slides',async(req,res)=>{try{
  const {lessonId,count=12,style='Hiện đại, rõ ràng',request='',source=''}=req.body; const ctx=sourceContext(lessonId,source); if(!ctx.lesson)return res.status(400).json({error:'Bài học không hợp lệ.'});
  const input=`BÀI HỌC:\n${JSON.stringify(ctx.lesson,null,2)}\n\nSỐ SLIDE: ${count}\nPHONG CÁCH: ${style}\nYÊU CẦU: ${request||'Không có'}\n\nKHO NGUỒN ĐÃ NẠP:\n${ctx.text||'(Không có nguồn chi tiết.)'}`;
  const data=await askAI(slideInstruction,input,slideSchema,'lesson_slides'); const project={id:id(),type:'slides',name:data.title||ctx.lesson.title,lessonId,createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),data}; const all=readProjects();all.unshift(project);writeJSON(projectsFile,all.slice(0,200));res.json({project});
}catch(e){res.status(500).json({error:e.message});}});

app.get('/api/projects',(_,res)=>res.json({projects:readProjects()}));
app.get('/api/projects/:id',(req,res)=>{const p=readProjects().find(x=>x.id===req.params.id);if(!p)return res.status(404).json({error:'Không tìm thấy sản phẩm.'});res.json({project:p});});
app.put('/api/projects/:id',(req,res)=>{const all=readProjects();const i=all.findIndex(x=>x.id===req.params.id);if(i<0)return res.status(404).json({error:'Không tìm thấy sản phẩm.'});const data=req.body.data;if(!data||typeof data!=='object')return res.status(400).json({error:'Dữ liệu chỉnh sửa không hợp lệ.'});all[i]={...all[i],data,updatedAt:new Date().toISOString(),name:data.title||all[i].name};writeJSON(projectsFile,all);res.json({project:all[i]});});
app.delete('/api/projects/:id',(req,res)=>{writeJSON(projectsFile,readProjects().filter(x=>x.id!==req.params.id));res.json({ok:true});});

const safeName=name=>clean(name).replace(/[^a-zA-Z0-9-_À-ỹ ]/g,'').slice(0,80).trim()||'edumath9';
app.post('/api/export/docx',async(req,res)=>{try{const p=req.body.project,d=p?.data;if(!d)return res.status(400).json({error:'Thiếu dữ liệu.'});const c=[new Paragraph({text:d.title||'KẾ HOẠCH BÀI DẠY',heading:HeadingLevel.TITLE}),new Paragraph({text:`Môn: ${d.subject||'Toán'} | Lớp: ${d.grade||'9'} | Thời lượng: ${d.duration||''}`}),new Paragraph({text:'I. MỤC TIÊU',heading:HeadingLevel.HEADING_1}),...(d.objectives||[]).map(x=>new Paragraph({text:x,bullet:{level:0}})),new Paragraph({text:'II. NĂNG LỰC VÀ PHẨM CHẤT',heading:HeadingLevel.HEADING_1}),...(d.competencies||[]).map(x=>new Paragraph({text:x,bullet:{level:0}})),...(d.qualities||[]).map(x=>new Paragraph({text:x,bullet:{level:0}})),new Paragraph({text:'III. CHUẨN BỊ',heading:HeadingLevel.HEADING_1}),new Paragraph({text:'Giáo viên: '+(d.preparation?.teacher||[]).join('; ')}),new Paragraph({text:'Học sinh: '+(d.preparation?.students||[]).join('; ')}),new Paragraph({text:'IV. TIẾN TRÌNH DẠY HỌC',heading:HeadingLevel.HEADING_1})];const rows=[new TableRow({children:['Giai đoạn','Thời gian','Hoạt động GV','Hoạt động HS','Sản phẩm/Đánh giá'].map(t=>new TableCell({children:[new Paragraph({text:t})]}))})];for(const a of d.activities||[])rows.push(new TableRow({children:[a.phase,a.time,a.teacher,a.students,`${a.product||''}\n${a.assessment||''}`].map(t=>new TableCell({children:[new Paragraph({text:String(t||'')})]}))}));c.push(new Table({rows,width:{size:100,type:WidthType.PERCENTAGE}}));c.push(new Paragraph({text:'V. KIẾN THỨC CẦN GHI NHỚ',heading:HeadingLevel.HEADING_1}),...(d.knowledgeSummary||[]).map(x=>new Paragraph({text:x,bullet:{level:0}})),new Paragraph({text:'VI. PHÂN HÓA',heading:HeadingLevel.HEADING_1}),new Paragraph({text:'Hỗ trợ: '+(d.differentiation?.support||[]).join('; ')}),new Paragraph({text:'Mở rộng: '+(d.differentiation?.extension||[]).join('; ')}),new Paragraph({text:'VII. DẶN DÒ',heading:HeadingLevel.HEADING_1}),...(d.homework||[]).map(x=>new Paragraph({text:x,bullet:{level:0}})));const buf=await Packer.toBuffer(new Document({sections:[{children:c}]}));const file=path.join(EXPORTS,`${safeName(d.title)}.docx`);fs.writeFileSync(file,buf);res.download(file,path.basename(file));}catch(e){res.status(500).json({error:e.message});}});
app.post('/api/export/pptx',async(req,res)=>{try{const p=req.body.project,d=p?.data;if(!d?.slides)return res.status(400).json({error:'Thiếu dữ liệu slide.'});const pptx=new pptxgen();pptx.layout='LAYOUT_WIDE';pptx.author='EduMath 9 AI';pptx.lang='vi-VN';pptx.title=d.title||'Bài giảng Toán 9';for(const [i,s] of d.slides.entries()){const sl=pptx.addSlide();sl.background={color:'F7FAFC'};sl.addText(s.title||'',{x:.65,y:.45,w:12,h:.7,fontSize:27,bold:true,margin:0});const bullets=(s.bullets||[]).map(x=>({text:String(x),options:{bullet:{indent:16},hanging:4}}));sl.addText(bullets,{x:.85,y:1.5,w:11.4,h:4.5,fontSize:19,breakLine:true,fit:'shrink',paraSpaceAfterPt:12});if(s.activity)sl.addText('Hoạt động: '+s.activity,{x:.85,y:6.15,w:11.4,h:.55,fontSize:13,italic:true,fit:'shrink'});if(s.speakerNotes)sl.addNotes(s.speakerNotes);sl.addText(`${i+1}/${d.slides.length}`,{x:11.8,y:7.0,w:.7,h:.2,fontSize:9,align:'right'});}const file=path.join(EXPORTS,`${safeName(d.title)}.pptx`);await pptx.writeFile({fileName:file});res.download(file,path.basename(file));}catch(e){res.status(500).json({error:e.message});}});

app.use((_,res)=>res.sendFile(path.join(ROOT,'public','index.html')));
const port=Number(process.env.PORT||3000);app.listen(port,()=>console.log(`EduMath 9 AI v2 running at http://localhost:${port}`));
