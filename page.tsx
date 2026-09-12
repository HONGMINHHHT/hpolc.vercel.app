"use client";
import {useMemo,useRef,useState} from "react";
import {AlertTriangle,BarChart3,Check,ChevronDown,Download,FileSpreadsheet,FileText,FolderOpen,GraduationCap,LayoutDashboard,Search,Settings,ShieldCheck,Sparkles,UploadCloud,Users,X} from "lucide-react";
import * as XLSX from "xlsx";

type Decision={id:string,date:string,type:string,students:number,groups:number[],warn?:boolean};
const data:Decision[]=[
 {id:"3401/QĐ-ĐHV",date:"12/12/2023",type:"Tuyển mới",students:29,groups:[10,0,5,5,9,0]},
 {id:"237/QĐ-ĐHV",date:"25/01/2024",type:"Tuyển bổ sung",students:7,groups:[1,0,1,1,4,0],warn:true},
 {id:"626/QĐ-ĐHV",date:"19/03/2024",type:"Tuyển bổ sung",students:20,groups:[4,0,2,0,3,11]},
 {id:"903/QĐ-ĐHV",date:"23/04/2024",type:"Tuyển bổ sung",students:4,groups:[0,0,0,0,0,4]},
 {id:"1382/QĐ-ĐHV",date:"05/06/2024",type:"Tuyển bổ sung",students:26,groups:[11,0,4,3,6,2]},
 {id:"1674/QĐ-ĐHV",date:"05/07/2024",type:"Tuyển bổ sung",students:27,groups:[4,0,3,6,4,10]},
];
const names=["ĐH (1)","CĐ (2)","TC (3)","CĐKN (4)","TCKN (5)","THPT (6)"];
const colors=["#1259a7","#7f8fa6","#12a37f","#f0a32f","#df6b34","#7b5cc7"];
const graduationMonths=[25,25,37,25,37,49];

function graduationDate(decisionDate:string,months:number){
 const [day,month,year]=decisionDate.split("/").map(Number);
 const result=new Date(Date.UTC(year,month-1,day));
 result.setUTCMonth(result.getUTCMonth()+months);
 return result;
}

export default function Home(){
 const [excel,setExcel]=useState<File|null>(null),[pdfs,setPdfs]=useState<File[]>([]),[busy,setBusy]=useState(false),[query,setQuery]=useState("");
 const xRef=useRef<HTMLInputElement>(null),pRef=useRef<HTMLInputElement>(null);
 const totals=useMemo(()=>data.reduce((a,d)=>a.map((v,i)=>v+d.groups[i]),[0,0,0,0,0,0]),[]);
 const filtered=data.filter(d=>d.id.toLowerCase().includes(query.toLowerCase()));
 function process(){if(!excel&&!pdfs.length)return;setBusy(true);setTimeout(()=>setBusy(false),1000)}
 function download(){
  const rows:(string|number|Date)[][]=[["TIẾN ĐỘ HỌC TẬP CỦA CÁC LỚP QUẢN LÝ"],["TT","CB Phụ trách","Ngành","Trạm","Lớp","QĐ trúng tuyển","Đối tượng","Số lượng","Số TC cần hoàn thành","Thời gian TN Dự kiến","Ghi chú"]];
  data.forEach((d,di)=>d.groups.forEach((n,gi)=>rows.push([gi?"":di+1,"","Luật","Trường Trung cấp Nghề Bình Thạnh","K64T1-Luật",gi?"":`${d.id} ngày ${d.date}\n${d.type}\nDanh sách tại trạm: ${d.students} SV`,names[gi],n,"",graduationDate(d.date,graduationMonths[gi]),d.warn&&gi===0?"Số tại Điều 1 khác danh sách của trạm/lớp":""])));
  rows.push(["","","","","","","TỔNG SỐ SINH VIÊN QUẢN LÝ",113,"","",""]);
  const wb=XLSX.utils.book_new();
  const summary=XLSX.utils.aoa_to_sheet(rows,{cellDates:true});summary["!cols"]=[{wch:5},{wch:18},{wch:14},{wch:38},{wch:16},{wch:45},{wch:16},{wch:12},{wch:20},{wch:23},{wch:42}];
  for(let row=3;row<=2+data.length*names.length;row++){const cell=summary[`J${row}`];if(cell)cell.z="dd/mm/yyyy";}
  XLSX.utils.book_append_sheet(wb,summary,"TONG_HOP_MAU");
  const detail=XLSX.utils.aoa_to_sheet([["STT","Họ và tên","Ngày sinh","Giới tính","Nơi sinh","Phân nhóm nguyên bản","Nhóm chuẩn hóa","Số Quyết định","Ngày Quyết định","Loại tuyển","Đợt tuyển","Ngành","Khóa","Lớp","Trạm/Cơ sở phối hợp","Tên file nguồn","Trang nguồn","Ghi chú kiểm tra"]]);
  XLSX.utils.book_append_sheet(wb,detail,"CHI_TIET_SINH_VIEN");
  const reconcile=XLSX.utils.aoa_to_sheet([["STT","Số QĐ","Ngày QĐ","Loại tuyển","Đợt","Ngành","Khóa","Lớp","Trạm chuẩn hóa","Số SV danh sách","ĐH","CĐ","TC","CĐKN","TCKN","THPT","Tổng kiểm tra","Chênh lệch","Ghi chú"],...data.map((d,i)=>[i+1,d.id,d.date,d.type,"","Luật",64,"K64T1-Luật","Trường Trung cấp Nghề Bình Thạnh",d.students,...d.groups,d.groups.reduce((a,b)=>a+b,0),"",d.warn?"Toàn quyết định và danh sách trạm/lớp có số lượng khác nhau":"Khớp dữ liệu"])]);
  XLSX.utils.book_append_sheet(wb,reconcile,"DOI_SOAT_QUYET_DINH");
  XLSX.writeFile(wb,"TONG_HOP_QUAN_LY_DAO_TAO_K64T1_LUAT.xlsx");
 }
 return <div className="min-h-screen bg-[#f3f7fb] text-[#17324d]">
  <header className="sticky top-0 z-30 border-b border-[#dce7f1] bg-white/95"><div className="mx-auto flex h-16 max-w-[1600px] items-center gap-4 px-4 lg:px-7">
   <div className="flex items-center gap-3"><div className="grid h-10 w-10 place-items-center rounded-xl bg-[#0e5cab] text-white shadow-lg"><GraduationCap size={23}/></div><div><div className="text-lg font-extrabold text-[#103f70]">HPOLC</div><div className="hidden text-xs text-[#70859a] sm:block">Hệ thống quản lý đào tạo</div></div></div>
   <label className="ml-auto hidden max-w-md flex-1 items-center rounded-xl border border-[#d9e4ee] bg-[#f7fafd] px-3 md:flex"><Search size={17} className="text-[#7790a7]"/><input value={query} onChange={e=>setQuery(e.target.value)} className="w-full bg-transparent px-2.5 py-2 text-sm outline-none" placeholder="Tìm số quyết định..."/></label>
   <button className="rounded-xl border border-[#dbe6ef] p-2.5 text-[#557087]" aria-label="Cài đặt"><Settings size={19}/></button><div className="grid h-9 w-9 place-items-center rounded-full bg-[#e4f0fb] text-sm font-bold text-[#0e5cab]">HP</div>
  </div></header>
  <div className="mx-auto grid max-w-[1600px] lg:grid-cols-[230px_1fr]">
   <aside className="hidden min-h-[calc(100vh-64px)] border-r border-[#dce7f1] bg-white p-4 lg:block"><nav className="space-y-1.5"><Nav active icon={<LayoutDashboard/>} text="Tổng quan"/><Nav icon={<UploadCloud/>} text="Nhập dữ liệu"/><Nav icon={<FileText/>} text="Quyết định" badge="6"/><Nav icon={<Users/>} text="Sinh viên" badge="113"/><Nav icon={<ShieldCheck/>} text="Đối soát" badge="1" warn/><Nav icon={<BarChart3/>} text="Báo cáo"/></nav><div className="mt-8 rounded-2xl bg-[#eef7ff] p-4"><div className="mb-2 flex items-center gap-2 text-sm font-bold text-[#0e5cab]"><Sparkles size={16}/> Chuẩn dữ liệu</div><p className="text-xs leading-5 text-[#617a91]">Kiểm tra theo danh sách kèm theo, không suy đoán từ tên tệp.</p></div></aside>
   <main className="min-w-0 p-4 sm:p-6 lg:p-8">
    <div className="mb-6 flex flex-wrap items-end justify-between gap-4"><div><p className="mb-1 text-sm font-bold text-[#0e73b9]">K64T1 · NGÀNH LUẬT</p><h1 className="text-2xl font-extrabold text-[#173f66] sm:text-3xl">Tổng hợp quản lý đào tạo</h1><p className="mt-1 text-sm text-[#6c8296]">Trường Trung cấp Nghề Bình Thạnh · Đào tạo từ xa</p></div><button onClick={download} className="flex items-center gap-2 rounded-xl bg-[#0e5cab] px-4 py-2.5 text-sm font-bold text-white shadow-lg"><Download size={18}/> Xuất bảng tổng hợp</button></div>
    <section className="mb-6 grid gap-4 xl:grid-cols-[1.25fr_.75fr]">
     <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="mb-4 flex items-center justify-between"><div><h2 className="font-bold text-[#244c70]">Nhập tài liệu</h2><p className="mt-0.5 text-sm text-[#71879b]">Excel mẫu và các PDF quyết định trúng tuyển</p></div><span className="rounded-full bg-[#edf6ff] px-3 py-1 text-xs font-bold text-[#0e66ad]">Bảo mật trong phiên</span></div>
      <div className="grid gap-3 sm:grid-cols-2"><UploadCard onClick={()=>xRef.current?.click()} icon={<FileSpreadsheet/>} tone="green" title={excel?.name||"Chọn Excel mẫu"} note=".xlsx hoặc .xls"/><UploadCard onClick={()=>pRef.current?.click()} icon={<FolderOpen/>} tone="orange" title={pdfs.length?`${pdfs.length} quyết định đã chọn`:"Chọn các quyết định PDF"} note="Có thể chọn nhiều tệp"/><input ref={xRef} hidden type="file" accept=".xlsx,.xls" onChange={e=>setExcel(e.target.files?.[0]||null)}/><input ref={pRef} hidden type="file" accept="application/pdf" multiple onChange={e=>setPdfs(Array.from(e.target.files||[]))}/></div>
      {(excel||pdfs.length>0)&&<div className="mt-3 flex flex-wrap gap-2 text-xs text-[#60798f]">{excel&&<Pill name={excel.name} remove={()=>setExcel(null)}/>} {pdfs.slice(0,2).map((f,i)=><Pill key={f.name+i} name={f.name} remove={()=>setPdfs(pdfs.filter((_,j)=>j!==i))}/>)} {pdfs.length>2&&<span>+{pdfs.length-2} tệp khác</span>}</div>}
      <div className="mt-4 flex items-center justify-between gap-3 border-t pt-4"><div><p className="text-xs leading-5 text-[#758a9c]">Giữ riêng số tại Điều 1 và số sinh viên trong danh sách kèm theo.</p><p className="text-xs font-semibold leading-5 text-[#0e68ae]">Thời gian TN dự kiến: ĐH/CĐ/CĐKN +25 tháng · TC/TCKN +37 tháng · THPT +49 tháng.</p></div><button onClick={process} disabled={busy||(!excel&&!pdfs.length)} className="shrink-0 rounded-xl bg-[#12a37f] px-4 py-2.5 text-sm font-bold text-white disabled:opacity-40">{busy?"Đang phân tích...":"Phân tích dữ liệu"}</button></div>
     </div>
     <div className="rounded-2xl bg-gradient-to-br from-[#0d579f] to-[#087cad] p-5 text-white shadow-xl"><div className="flex items-start justify-between"><div><p className="text-sm font-semibold text-[#cceaff]">TỔNG SỐ SINH VIÊN</p><p className="mt-2 text-5xl font-extrabold">113</p></div><span className="grid h-12 w-12 place-items-center rounded-2xl bg-white/15"><Users/></span></div><div className="mt-7 h-2 rounded-full bg-white/15"><div className="h-full w-full rounded-full bg-[#53ddae]"/></div><div className="mt-3 flex justify-between text-sm"><span>6/6 quyết định</span><span className="flex items-center gap-1 font-bold"><Check size={16}/> Đủ dữ liệu</span></div><div className="mt-5 grid grid-cols-3 border-t border-white/20 pt-4 text-center"><Mini n="30" text="Đại học"/><Mini n="26" text="TCKN"/><Mini n="27" text="THPT"/></div></div>
    </section>
    <section className="mb-6 grid grid-cols-2 gap-3 sm:grid-cols-3 xl:grid-cols-6">{names.map((n,i)=><div key={n} className="rounded-2xl border bg-white p-4"><div className="mb-3 h-1.5 w-9 rounded-full" style={{background:colors[i]}}/><div className="text-2xl font-extrabold text-[#234a6d]">{totals[i]}</div><div className="mt-1 text-sm font-semibold text-[#71879b]">{n}</div></div>)}</section>
    <section className="overflow-hidden rounded-2xl border bg-white shadow-sm"><div className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-4"><div><h2 className="font-bold text-[#244c70]">Đối soát quyết định</h2><p className="mt-0.5 text-sm text-[#778b9c]">Tổng nhóm đối tượng phải bằng danh sách kèm theo</p></div><button className="flex items-center gap-2 rounded-lg border px-3 py-2 text-sm font-semibold text-[#587189]">Tất cả trạng thái <ChevronDown size={15}/></button></div>
     <div className="overflow-x-auto"><table className="w-full min-w-[820px] text-left text-sm"><thead className="bg-[#f5f9fc] text-xs uppercase text-[#668097]"><tr><th className="px-5 py-3.5">Quyết định</th><th className="px-4">Ngày</th><th className="px-4">Loại tuyển</th><th className="px-4 text-center">Danh sách</th><th className="px-4 text-center">Tổng kiểm tra</th><th className="px-4">Kết quả</th><th/></tr></thead><tbody className="divide-y">{filtered.map(d=><tr key={d.id} className="hover:bg-[#f8fbfd]"><td className="px-5 py-4 font-bold text-[#225985]">{d.id}</td><td className="px-4 text-[#5f788e]">{d.date}</td><td className="px-4 text-[#5f788e]">{d.type}</td><td className="px-4 text-center font-bold">{d.students} SV</td><td className="px-4 text-center font-bold">{d.groups.reduce((a,b)=>a+b,0)} SV</td><td className="px-4">{d.warn?<Status warn/>:<Status/>}</td><td className="px-5 text-right"><button className="font-bold text-[#0e68ae]">Chi tiết</button></td></tr>)}</tbody></table></div>
     <div className="flex flex-wrap items-center justify-between gap-3 border-t bg-[#fbfdff] px-5 py-4"><p className="text-sm text-[#688196]">Đã kiểm tra <b>6 quyết định</b> · Không có bản ghi trùng</p><div className="flex items-center gap-2 rounded-lg bg-[#fff5e7] px-3 py-2 text-xs font-semibold text-[#9b610d]"><AlertTriangle size={15}/> QĐ 237: Điều 1 khác danh sách của trạm</div></div>
    </section>
   </main>
  </div>
 </div>
}
function Nav({icon,text,active,badge,warn}:{icon:React.ReactElement,text:string,active?:boolean,badge?:string,warn?:boolean}){return <button className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-semibold ${active?"bg-[#eaf4fd] text-[#0e5cab]":"text-[#637b90] hover:bg-[#f4f8fb]"}`}>{useMemo(()=>icon,[])}<span>{text}</span>{badge&&<span className={`ml-auto rounded-full px-2 py-.5 text-xs ${warn?"bg-[#fff0dc] text-[#b36809]":"bg-[#eef3f7]"}`}>{badge}</span>}</button>}
function UploadCard({onClick,icon,title,note,tone}:{onClick:()=>void,icon:React.ReactNode,title:string,note:string,tone:string}){return <button onClick={onClick} className="flex min-h-28 items-center gap-4 rounded-xl border border-dashed border-[#b9cddd] bg-[#f8fbfe] p-4 text-left hover:border-[#0e73b9]"><span className={`grid h-11 w-11 place-items-center rounded-xl ${tone==="green"?"bg-[#e7f5ed] text-[#138452]":"bg-[#fff0eb] text-[#df5f32]"}`}>{icon}</span><span className="min-w-0"><b className="block truncate text-sm text-[#284d6d]">{title}</b><span className="mt-1 block text-xs text-[#7890a4]">{note}</span></span></button>}
function Pill({name,remove}:{name:string,remove:()=>void}){return <span className="inline-flex max-w-60 items-center gap-1 rounded-full bg-[#eef4f8] px-2.5 py-1"><span className="truncate">{name}</span><button onClick={remove}><X size={13}/></button></span>}
function Mini({n,text}:{n:string,text:string}){return <div><div className="text-xl font-extrabold">{n}</div><div className="text-xs text-[#cce7f6]">{text}</div></div>}
function Status({warn}:{warn?:boolean}){return warn?<span className="inline-flex items-center gap-1 rounded-full bg-[#fff3df] px-2.5 py-1 text-xs font-bold text-[#a96206]"><AlertTriangle size={13}/>Có lưu ý</span>:<span className="inline-flex items-center gap-1 rounded-full bg-[#e8f8f1] px-2.5 py-1 text-xs font-bold text-[#087a59]"><Check size={13}/>Khớp dữ liệu</span>}
