const ATTENDANCE_SICK_LEAVE_STATUS = 'اجازة مرضية';
const ATTENDANCE_MEDICAL_FILE_MAX_BYTES = 5 * 1024 * 1024;

function fileToBase64_(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || '').split(',')[1] || '');
        reader.onerror = () => reject(new Error('تعذرت قراءة ملف التقرير الطبي'));
        reader.readAsDataURL(file);
    });
}

async function handleAttendancePage() {
    const user = JSON.parse(localStorage.getItem('currentUser')) || JSON.parse(sessionStorage.getItem('currentUser'));
    const form=document.getElementById('attendanceForm'), body=document.getElementById('attendance-table-body');
    const role=String(user.role||'').toLowerCase(), statusFilter=document.getElementById('attendanceFilterStatus');
    const employeeFilter=document.getElementById('attendanceEmployeeFilter'), message=document.getElementById('attendanceStatusMessage');
    const statusSelect=document.getElementById('attendanceStatus'), medicalWrap=document.getElementById('medicalReportWrap'), medicalFile=document.getElementById('medicalReportFile');
    let rows=[];
    const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
    const canReview = role === 'admin' || role === 'manager' || role === 'auditor';
    if (canReview) document.getElementById('attendanceReviewHead')?.classList.remove('d-none');
    const reviewAttendance = async (button, timestamp, username, status) => {
        try {
            const rejectReason = status === 'rejected' ? (prompt('سبب رفض سجل الدوام:') || '').trim() : '';
            if (status === 'rejected' && !rejectReason) { alert('يرجى كتابة سبب الرفض.'); return; }
            if (status === 'approved' && !confirm('اعتماد سجل الدوام هذا؟')) return;
            button.disabled = true;
            const original = button.innerHTML;
            button.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
            const res = await fetch(SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'approveAttendance', payload: { timestamp, username, role: user.role || '', status, reason: rejectReason, approvedBy: user.name || '', approvedAt: new Date().toISOString() } })
            });
            const result = await res.json();
            if (!result || result.status !== 'success') throw new Error(result?.message || 'فشل العملية');
            localStorage.removeItem(attendanceCacheKey());
            await load();
            message.className = 'alert alert-success mt-3';
            message.textContent = status === 'approved' ? 'تم اعتماد سجل الدوام.' : 'تم رفض سجل الدوام.';
        } catch (e) {
            button.disabled = false;
            button.innerHTML = original;
            alert(`تعذر ${status === 'approved' ? 'اعتماد' : 'رفض'} سجل الدوام: ${e.message || e}\n\nتأكد من نشر النسخة المحدّثة من Apps Script.`);
        }
    };
    const bindAttendanceReview = () => {
        body.querySelectorAll('.approve-att-btn').forEach(b => b.addEventListener('click', () => reviewAttendance(b, b.dataset.ts, b.dataset.u, 'approved')));
        body.querySelectorAll('.reject-att-btn').forEach(b => b.addEventListener('click', () => reviewAttendance(b, b.dataset.ts, b.dataset.u, 'rejected')));
    };
    const render=()=>{const filtered=rows.filter(r=>!statusFilter.value||r.status===statusFilter.value);document.getElementById('attendanceCount').textContent=`${filtered.length} سجل`;
    const emptyCols = canReview ? 7 : 6;
    body.innerHTML=filtered.length?filtered.map(r=>{
        const review = canReview ? (
            `<td class="text-nowrap">${r.approvalStatus === 'approved' ? '<span class="badge bg-success">معتمد</span>' : (r.approvalStatus === 'rejected' ? '<span class="badge bg-danger">مرفوض</span>' : '<span class="badge bg-secondary">قيد المراجعة</span>')} ` +
            (r.approvalStatus !== 'approved' ? `<button class="btn btn-sm btn-outline-success approve-att-btn" data-ts="${esc(r.timestamp)}" data-u="${esc(r.username)}" title="اعتماد السجل"><i class="fa-solid fa-check"></i></button> ` : '') +
            (r.approvalStatus !== 'rejected' ? `<button class="btn btn-sm btn-outline-danger reject-att-btn" data-ts="${esc(r.timestamp)}" data-u="${esc(r.username)}" title="رفض السجل"><i class="fa-solid fa-ban"></i></button>` : '') + '</td>'
        ) : '';
        return `<tr><td dir="ltr">${esc(r.timestamp)}</td><td>${esc(r.username)}</td><td><span class="badge ${r.status==='بداية دوام'?'bg-success':'bg-secondary'}">${esc(r.status)}</span></td><td>${esc(r.attendanceStatement)}</td><td>${esc(r.notes)||'-'}</td><td>${r.medicalReportUrl?`<a href="${esc(r.medicalReportUrl)}" target="_blank" rel="noopener" class="btn btn-sm btn-outline-primary"><i class="fa-solid fa-file-medical me-1"></i>عرض</a>`:'-'}</td>${review}</tr>`;
    }).join(''):`<tr><td colspan="${emptyCols}" class="text-center text-muted py-4">لا توجد سجلات مطابقة</td></tr>`;bindAttendanceReview();};

    // ---- Smart cache-first loading (instant render from localStorage, refresh
    // in the background, and keep showing the cached list if offline/failed). ----
    const attendanceCacheKey=()=>`attendanceCache::${employeeFilter?.value||'all'}`;
    const readAttendanceCache=()=>{try{const raw=localStorage.getItem(attendanceCacheKey());if(!raw)return null;const parsed=JSON.parse(raw);return Array.isArray(parsed)?parsed:null;}catch(e){return null;}};
    const writeAttendanceCache=(data)=>{try{localStorage.setItem(attendanceCacheKey(),JSON.stringify(data));}catch(e){/* تجاهل امتلاء التخزين المحلي */}};
    const load=async()=>{
        const cached=readAttendanceCache();
        if(cached){rows=cached;render();} else {body.innerHTML='<tr><td colspan="6" class="text-center py-4">جار التحميل...</td></tr>';}
        try{
            const r=await apiGet('getAttendance',{userId:String(user.id||''),role:String(user.role||''),userName:String(user.username||user.name||''),targetUserId:employeeFilter?.value||''});
            if(!Array.isArray(r))throw Error(r.message||'تعذر تحميل السجل');
            rows=r;render();writeAttendanceCache(rows);
        }catch(e){
            if(!cached) body.innerHTML=`<tr><td colspan="6" class="text-center text-danger">${esc(e.message)}</td></tr>`;
            // إذا فيه كاش، منخليه ظاهر كما هو (أوفلاين) بدل ما نستبدله برسالة خطأ.
        }
    };
    window.addEventListener('attendanceCacheInvalidated', load);

    // Status options are pulled live from the "statusWT" sheet (auto-created
    // with defaults on first run by the backend) so the dropdown (form) and
    // the history filter always match what's configured there. Cached locally
    // too, so the dropdown appears instantly and still works offline.
    const STATUS_CACHE_KEY='attendanceStatusCache';
    const readStatusCache=()=>{try{const raw=localStorage.getItem(STATUS_CACHE_KEY);const parsed=raw?JSON.parse(raw):null;return Array.isArray(parsed)&&parsed.length?parsed:null;}catch(e){return null;}};
    const writeStatusCache=(options)=>{try{localStorage.setItem(STATUS_CACHE_KEY,JSON.stringify(options));}catch(e){}};
    const applyStatusOptions=(options)=>{
        const currentValue=statusSelect.value;
        statusSelect.innerHTML=options.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('');
        if(options.includes(currentValue)) statusSelect.value=currentValue;
        const currentFilter=statusFilter.value;
        statusFilter.innerHTML='<option value="">كل الحالات</option>'+options.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('');
        if(options.includes(currentFilter)) statusFilter.value=currentFilter;
        toggleMedicalReportField();
    };
    const loadStatusOptions=async()=>{
        const cached=readStatusCache();
        if(cached) applyStatusOptions(cached);
        try{
            const r=await apiGet('getStatusOptions');
            if(r&&r.status==='success'&&Array.isArray(r.options)&&r.options.length){applyStatusOptions(r.options);writeStatusCache(r.options);return;}
        }catch(e){console.warn('تعذر تحميل حالات الدوام من statusWT (سيتم استخدام القيم المخزنة/الافتراضية)',e);}
        if(!cached) applyStatusOptions(['بداية دوام','نهاية دوام','عطلة أسبوعية','عطلة رسمية','اجازة إدارية','اجازة مرضية','حضور إضافي']);
    };
    const toggleMedicalReportField=()=>{
        const isSick=statusSelect.value===ATTENDANCE_SICK_LEAVE_STATUS;
        medicalWrap.classList.toggle('d-none',!isSick);
        medicalFile.required=isSick;
        if(!isSick) medicalFile.value='';
    };

    if(role==='admin'||role==='manager'){try{const options=await fetchTeamOptions(user);employeeFilter.innerHTML='<option value="">كل الموظفين المسموح بهم</option>'+options.map(o=>`<option value="${esc(o.id)}">${esc(o.name)}</option>`).join('');document.getElementById('attendanceEmployeeFilterWrap').classList.remove('d-none');}catch(e){console.warn(e);}}
    statusFilter.addEventListener('change',render);employeeFilter?.addEventListener('change',load);
    statusSelect.addEventListener('change',toggleMedicalReportField);
    await loadStatusOptions();

    form.addEventListener('submit',async e=>{
        e.preventDefault();
        const b=form.querySelector('button[type=submit]');
        b.disabled=true;
        try{
            const status=statusSelect.value;
            const payload={username:user.username||user.name,status,attendanceStatement:document.getElementById('attendanceStatement').value.trim(),notes:document.getElementById('attendanceNotes').value.trim()};
            if(status===ATTENDANCE_SICK_LEAVE_STATUS){
                const file=medicalFile.files && medicalFile.files[0];
                if(!file) throw new Error('يجب إرفاق التقرير الطبي (PDF أو صورة) عند اختيار حالة اجازة مرضية');
                if(file.size>ATTENDANCE_MEDICAL_FILE_MAX_BYTES) throw new Error('حجم ملف التقرير الطبي يتجاوز 5 ميجابايت');
                payload.medicalReportBase64=await fileToBase64_(file);
                payload.medicalReportFileName=file.name;
                payload.medicalReportMimeType=file.type||'application/octet-stream';
            }

            const saveOnlineOrQueue=async()=>{
                if(!navigator.onLine){await queueAttendanceOffline(payload);return{queued:true};}
                try{
                    const r=await(await fetch(SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'submitAttendance',payload})})).json();
                    if(r.status!=='success')throw Error(r.message||'فشل الحفظ');
                    return r;
                }catch(err){
                    const networkFailure=!navigator.onLine||err instanceof TypeError||/failed to fetch|network|load failed/i.test(String(err.message||''));
                    if(networkFailure){await queueAttendanceOffline(payload);return{queued:true};}
                    throw err;
                }
            };

            const result=await saveOnlineOrQueue();
            form.reset();
            toggleMedicalReportField();
            if(result.queued){
                updateOfflineStatus();
                message.className='alert alert-warning mt-3';
                message.textContent='تم حفظ السجل محلياً وسيتم إرساله تلقائياً عند عودة الإنترنت.';
            }else{
                message.className='alert alert-success mt-3';
                message.textContent='تم تسجيل الدوام بنجاح.';
                localStorage.removeItem(attendanceCacheKey());
                await load();
            }
        }catch(err){
            message.className='alert alert-danger mt-3';
            message.textContent=err.message;
        }finally{
            b.disabled=false;
            b.innerHTML='<i class="fa-solid fa-check me-1"></i> تسجيل الدوام';
        }
    });
    // V58: عند إعادة زيارة شاشة الدوام من قائمة التنقل نُعيد تحميل السجلات.
    window.addEventListener('spaViewRevisited', async (event) => {
        if (event.detail && event.detail.route !== 'attendance') return;
        await load();
    });

    await load();
}

// V58: تسجيل شاشة الدوام في موجه الـ SPA.
if (typeof registerView === 'function') registerView('attendance', handleAttendancePage);
