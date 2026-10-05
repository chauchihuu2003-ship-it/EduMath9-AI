import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI } from '@google/genai';
import mammoth from 'mammoth';
import { PDFParse } from 'pdf-parse';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ShadingType, VerticalAlign, Footer, PageNumber } from 'docx';
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

const client = process.env.GEMINI_API_KEY ? new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY }) : null;
const model = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';
const clean = s => String(s ?? '').replace(/\u0000/g, '').trim();
const clamp = (s, n = 50000) => clean(s).slice(0, n);
const ensureAI = () => { if (!client) throw new Error('Chưa cấu hình GEMINI_API_KEY trong file .env'); };

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
  const prompt = `${instructions}\n\n${input}\n\nYÊU CẦU ĐỊNH DẠNG: Trả về đúng một JSON object theo schema được cung cấp. Không thêm markdown, không thêm lời giải thích.`;
  const r = await client.models.generateContent({
    model,
    contents: prompt,
    config: {
      responseMimeType: 'application/json',
      responseSchema: schema,
      temperature: 0.4
    }
  });
  const text = r.text?.trim();
  if (!text) throw new Error('Gemini không trả về nội dung.');
  try { return JSON.parse(text); }
  catch { throw new Error('Gemini trả về JSON không hợp lệ.'); }
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

app.get('/api/health', (_,res)=>res.json({ok:true,aiConfigured:!!client,provider:'gemini',model,version:'3.0'}));
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
app.post('/api/export/docx',async(req,res)=>{try{
  const p=req.body.project,d=p?.data;
  if(!d)return res.status(400).json({error:'Thiếu dữ liệu.'});
  const font='Times New Roman';
  const run=(text='',opts={})=>new TextRun({text:String(text??''),font,size:opts.size||26,bold:!!opts.bold,italics:!!opts.italics});
  const para=(text='',opts={})=>new Paragraph({alignment:opts.alignment,style:opts.style,spacing:{line:360,after:opts.after??100,before:opts.before??0},indent:opts.indent,children:[run(text,{size:opts.size,bold:opts.bold,italics:opts.italics})]});
  const bullet=(text)=>new Paragraph({bullet:{level:0},spacing:{line:360,after:70},children:[run(text)]});
  const heading=(text)=>new Paragraph({spacing:{before:240,after:100,line:360},children:[run(text,{size:28,bold:true})]});
  const centerTitle=new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100,line:360},children:[run(d.title||'KẾ HOẠCH BÀI DẠY',{size:32,bold:true})]});
  const meta=new Table({width:{size:100,type:WidthType.PERCENTAGE},borders:{top:{style:BorderStyle.SINGLE,size:6,color:'B7B7B7'},bottom:{style:BorderStyle.SINGLE,size:6,color:'B7B7B7'},left:{style:BorderStyle.SINGLE,size:6,color:'B7B7B7'},right:{style:BorderStyle.SINGLE,size:6,color:'B7B7B7'},insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:'D9D9D9'},insideVertical:{style:BorderStyle.SINGLE,size:4,color:'D9D9D9'}},rows:[new TableRow({children:[
    new TableCell({width:{size:25,type:WidthType.PERCENTAGE},children:[para('Môn học',{bold:true})]}),new TableCell({width:{size:25,type:WidthType.PERCENTAGE},children:[para(d.subject||'Toán')]}),
    new TableCell({width:{size:25,type:WidthType.PERCENTAGE},children:[para('Lớp',{bold:true})]}),new TableCell({width:{size:25,type:WidthType.PERCENTAGE},children:[para(d.grade||'9')]})
  ]}),new TableRow({children:[
    new TableCell({children:[para('Thời lượng',{bold:true})]}),new TableCell({children:[para(d.duration||'')] }),
    new TableCell({children:[para('Bộ sách',{bold:true})]}),new TableCell({children:[para('Kết nối tri thức với cuộc sống')]})
  ]})]});
  const c=[centerTitle,meta,heading('I. MỤC TIÊU')];
  for(const x of d.objectives||[])c.push(bullet(x));
  c.push(heading('II. NĂNG LỰC VÀ PHẨM CHẤT'),new Paragraph({spacing:{before:80,after:60},children:[run('Năng lực',{bold:true})]}));
  for(const x of d.competencies||[])c.push(bullet(x));
  c.push(new Paragraph({spacing:{before:80,after:60},children:[run('Phẩm chất',{bold:true})]}));
  for(const x of d.qualities||[])c.push(bullet(x));
  c.push(heading('III. CHUẨN BỊ'),new Paragraph({spacing:{after:60},children:[run('1. Giáo viên',{bold:true})]}));
  for(const x of d.preparation?.teacher||[])c.push(bullet(x));
  c.push(new Paragraph({spacing:{before:80,after:60},children:[run('2. Học sinh',{bold:true})]}));
  for(const x of d.preparation?.students||[])c.push(bullet(x));
  c.push(heading('IV. TIẾN TRÌNH DẠY HỌC'));
  const border={top:{style:BorderStyle.SINGLE,size:6,color:'333333'},bottom:{style:BorderStyle.SINGLE,size:6,color:'333333'},left:{style:BorderStyle.SINGLE,size:6,color:'333333'},right:{style:BorderStyle.SINGLE,size:6,color:'333333'},insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:'555555'},insideVertical:{style:BorderStyle.SINGLE,size:4,color:'555555'}};
  const cell=(text,bold=false)=>new TableCell({verticalAlign:VerticalAlign.TOP,children:[new Paragraph({spacing:{line:330,after:40},children:[run(text||'',{size:22,bold})]})]});
  const rows=[new TableRow({children:[cell('Giai đoạn',true),cell('Thời gian',true),cell('Hoạt động GV',true),cell('Hoạt động HS',true),cell('Sản phẩm / Đánh giá',true)]})];
  for(const a of d.activities||[])rows.push(new TableRow({children:[cell(a.phase),cell(a.time),cell(a.teacher),cell(a.students),cell(`${a.product||''}${a.assessment?'\nĐánh giá: '+a.assessment:''}`)]}));
  c.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},borders:border,rows}));
  c.push(heading('V. KIẾN THỨC CẦN GHI NHỚ'));for(const x of d.knowledgeSummary||[])c.push(bullet(x));
  c.push(heading('VI. PHÂN HÓA'),new Paragraph({spacing:{after:60},children:[run('Hỗ trợ',{bold:true})]}));for(const x of d.differentiation?.support||[])c.push(bullet(x));
  c.push(new Paragraph({spacing:{before:80,after:60},children:[run('Mở rộng',{bold:true})]}));for(const x of d.differentiation?.extension||[])c.push(bullet(x));
  c.push(heading('VII. DẶN DÒ'));for(const x of d.homework||[])c.push(bullet(x));
  if((d.notes||[]).length){c.push(heading('VIII. GHI CHÚ'));for(const x of d.notes||[])c.push(bullet(x));}
  c.push(new Paragraph({spacing:{before:300,after:0},children:[run('Bản nháp được tạo bằng EduMath 9 AI. Giáo viên cần rà soát nội dung chuyên môn trước khi sử dụng.',{size:20,italics:true})]}));
  const doc=new Document({styles:{default:{document:{run:{font,size:26},paragraph:{spacing:{line:360,after:100}}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1417,right:1417,bottom:1417,left:1417}}},footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[run('EduMath 9 AI · ',{size:18,italics:true}),new PageNumber({})]})]})},children:c}]});
  const buf=await Packer.toBuffer(doc);const file=path.join(EXPORTS,`${safeName(d.title)}.docx`);fs.writeFileSync(file,buf);res.download(file,path.basename(file));
}catch(e){res.status(500).json({error:e.message});}});
app.post('/api/export/pptx',async(req,res)=>{try{const p=req.body.project,d=p?.data;if(!d?.slides)return res.status(400).json({error:'Thiếu dữ liệu slide.'});const pptx=new pptxgen();pptx.layout='LAYOUT_WIDE';pptx.author='EduMath 9 AI';pptx.lang='vi-VN';pptx.title=d.title||'Bài giảng Toán 9';for(const [i,s] of d.slides.entries()){const sl=pptx.addSlide();sl.background={color:'F7FAFC'};sl.addText(s.title||'',{x:.65,y:.45,w:12,h:.7,fontSize:27,bold:true,margin:0});const bullets=(s.bullets||[]).map(x=>({text:String(x),options:{bullet:{indent:16},hanging:4}}));sl.addText(bullets,{x:.85,y:1.5,w:11.4,h:4.5,fontSize:19,breakLine:true,fit:'shrink',paraSpaceAfterPt:12});if(s.activity)sl.addText('Hoạt động: '+s.activity,{x:.85,y:6.15,w:11.4,h:.55,fontSize:13,italic:true,fit:'shrink'});if(s.speakerNotes)sl.addNotes(s.speakerNotes);sl.addText(`${i+1}/${d.slides.length}`,{x:11.8,y:7.0,w:.7,h:.2,fontSize:9,align:'right'});}const file=path.join(EXPORTS,`${safeName(d.title)}.pptx`);await pptx.writeFile({fileName:file});res.download(file,path.basename(file));}catch(e){res.status(500).json({error:e.message});}});

app.use((_,res)=>res.sendFile(path.join(ROOT,'public','index.html')));
const port=Number(process.env.PORT||3000);app.listen(port,()=>console.log(`EduMath 9 AI v2 running at http://localhost:${port}`));
