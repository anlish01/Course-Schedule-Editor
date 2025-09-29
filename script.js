class TimetableApp {
    constructor() {
        this.subjects = [
            { id: '1', name: '语文', teacher: '', color: '#FF0000' },
            { id: '2', name: '数学', teacher: '', color: '#0000FF' },
            { id: '3', name: '英语', teacher: '', color: '#FF69B4' },
            { id: '4', name: '美术', teacher: '', color: '#FFA500' },
            { id: '5', name: '健康与体育', teacher: '', color: '#800080' },
            { id: '6', name: '道德与法制', teacher: '', color: '#B8860B' },
            { id: '7', name: '音乐', teacher: '', color: '#FF4500' },
            { id: '8', name: '劳动', teacher: '', color: '#808080' },
            { id: '9', name: '班队会', teacher: '', color: '#663399' },
            { id: '10', name: '值日', teacher: '', color: '#006400' },
            { id: '11', name: '历史', teacher: '', color: '#000080' },
            { id: '12', name: '地理', teacher: '', color: '#4B0082' },
            { id: '13', name: '政治', teacher: '', color: '#8B0000' },
            { id: '14', name: '物理', teacher: '', color: '#A0522D' },
            { id: '15', name: '化学', teacher: '', color: '#2F4F4F' },
            { id: '16', name: '生物', teacher: '', color: '#FF6B6B' }
        ];
        this.timetable = {};
        this.periods = {
            morning: [
                { name: '第1节', time: '08:00-08:40' },
                { name: '第2节', time: '08:50-09:30' },
                { name: '第3节', time: '10:00-10:40' },
                { name: '第4节', time: '10:50-11:30' }
            ],
            afternoon: [
                { name: '第1节', time: '14:00-14:40' },
                { name: '第2节', time: '14:50-15:30' },
                { name: '第3节', time: '15:40-16:20' }
            ],
            evening: [
                { name: '第1节', time: '19:00-19:40' },
                { name: '第2节', time: '19:50-20:30' }
            ]
        };
        this.settings = {
            showEvening: true,
            showSaturday: true,
            showSunday: true,
            showPeriodTime: true
        };
        this.editingSubject = null;
        this.editingCell = null;
        this.editingPeriod = null;
        this.draggedSubject = null;
        
        this.init();
    }

    init() {
        this.loadData();
        this.loadSettings();
        this.bindEvents();
        this.renderSubjects();
        this.renderTimetable();
        this.loadTimetableTitle();
        this.loadTableTitle();
        this.applySettings();
    }

    bindEvents() {
        // 科目相关
        document.getElementById('addSubjectBtn').addEventListener('click', () => this.openSubjectModal());
        document.getElementById('subjectForm').addEventListener('submit', (e) => this.saveSubject(e));
        document.getElementById('cancelBtn').addEventListener('click', () => this.closeSubjectModal());
        document.getElementById('deleteSubjectBtn').addEventListener('click', () => this.deleteSubject());
        
        // 课程表标题
        document.getElementById('timetableTitle').addEventListener('input', (e) => this.saveTimetableTitle(e.target.value));
        document.getElementById('tableTitle').addEventListener('input', (e) => this.saveTableTitle(e.target.value));
        
        // 课时管理
        document.getElementById('addMorningBtn').addEventListener('click', () => this.addPeriod('morning'));
        document.getElementById('addAfternoonBtn').addEventListener('click', () => this.addPeriod('afternoon'));
        document.getElementById('addEveningBtn').addEventListener('click', () => this.addPeriod('evening'));
        document.getElementById('removeMorningBtn').addEventListener('click', () => this.removePeriod('morning'));
        document.getElementById('removeAfternoonBtn').addEventListener('click', () => this.removePeriod('afternoon'));
        document.getElementById('removeEveningBtn').addEventListener('click', () => this.removePeriod('evening'));
        
        // PC端课时管理按钮
        document.getElementById('addMorningBtn2').addEventListener('click', () => this.addPeriod('morning'));
        document.getElementById('addAfternoonBtn2').addEventListener('click', () => this.addPeriod('afternoon'));
        document.getElementById('addEveningBtn2').addEventListener('click', () => this.addPeriod('evening'));
        document.getElementById('removeMorningBtn2').addEventListener('click', () => this.removePeriod('morning'));
        document.getElementById('removeAfternoonBtn2').addEventListener('click', () => this.removePeriod('afternoon'));
        document.getElementById('removeEveningBtn2').addEventListener('click', () => this.removePeriod('evening'));
        
        // 时间相关
        document.getElementById('timeForm').addEventListener('submit', (e) => this.savePeriodTime(e));
        document.getElementById('cancelTimeBtn').addEventListener('click', () => this.closeTimeModal());
        
        // 初始化时间选择器
        this.initTimeSelectors();
        
        // 教程事件
        document.getElementById('tutorialBtn').addEventListener('click', () => this.openTutorialModal());
        document.getElementById('closeTutorialBtn').addEventListener('click', () => this.closeTutorialModal());
        
        // 重置和导出
        document.getElementById('resetBtn').addEventListener('click', () => this.resetTimetable());
        
        // 导出下拉菜单
        document.getElementById('exportBtn').addEventListener('click', (e) => this.toggleExportDropdown(e));
        document.getElementById('saveImageBtn').addEventListener('click', () => this.saveAsImage());
        document.getElementById('exportWordBtn').addEventListener('click', () => this.exportToWord());
        document.getElementById('exportExcelBtn').addEventListener('click', () => this.exportToExcel());
        
        // 设置相关
        document.getElementById('settingsBtn').addEventListener('click', () => this.openSettingsModal());
        document.getElementById('settingsForm').addEventListener('submit', (e) => this.saveSettings(e));
        document.getElementById('cancelSettingsBtn').addEventListener('click', () => this.closeSettingsModal());
        
        // 备份数据相关
        document.getElementById('backupBtn').addEventListener('click', (e) => this.toggleBackupMenu(e));
        document.getElementById('exportDataBtn').addEventListener('click', () => this.exportData());
        document.getElementById('importDataBtn').addEventListener('click', () => this.importData());
        document.getElementById('importFileInput').addEventListener('change', (e) => this.handleFileImport(e));
        
        // 手机端汉堡菜单相关
        document.getElementById('hamburgerBtn').addEventListener('click', () => this.toggleMobileSidebar());
        document.getElementById('closeSidebarBtn').addEventListener('click', () => this.closeMobileSidebar());
        document.getElementById('sidebarOverlay').addEventListener('click', () => this.closeMobileSidebar());
        
        // 侧边栏菜单按钮事件
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('sidebar-btn')) {
                const action = e.target.dataset.action;
                this.handleSidebarAction(action);
            }
        });
        
        // 全局点击事件监听器 - 点击外部区域隐藏下拉菜单
        document.addEventListener('click', (e) => this.handleGlobalClick(e));
        
        // 窗口大小改变时重新调整下拉菜单
        window.addEventListener('resize', () => this.handleWindowResize());
        

        
        // 颜色选择
        document.querySelectorAll('.color-option').forEach(option => {
            option.addEventListener('click', (e) => this.selectColor(e));
        });
        document.getElementById('customColor').addEventListener('change', (e) => {
            document.getElementById('customColorText').value = e.target.value;
        });
        document.getElementById('customColorText').addEventListener('input', (e) => {
            if (/^#[0-9A-F]{6}$/i.test(e.target.value)) {
                document.getElementById('customColor').value = e.target.value;
            }
        });
        
        // 拖拽相关
        this.setupDragAndDrop();
        
        // 键盘事件
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Delete' && this.editingCell) {
                this.removeSubjectFromCell(this.editingCell);
            }
            if (e.key === 'Escape') {
                // 关闭所有弹窗（优先级：手机端侧边栏 > 手机端科目选择 > 科目编辑 > 时间设置 > 教程）
                const sidebar = document.getElementById('mobileSidebar');
                if (sidebar && sidebar.classList.contains('show')) {
                    this.closeMobileSidebar();
                } else if (this.currentMobileModal) {
                    this.closeMobileSubjectModal(this.currentMobileModal);
                } else {
                    this.closeSubjectModal();
                    this.closeTimeModal();
                    this.closeTutorialModal();
                }
            }
        });
    }

    setupDragAndDrop() {
        // 科目池拖拽
        document.getElementById('subjectPool').addEventListener('dragstart', (e) => {
            if (e.target.classList.contains('subject-card')) {
                this.draggedSubject = e.target.dataset.subjectId;
                e.target.classList.add('dragging');
            }
        });
        
        document.getElementById('subjectPool').addEventListener('dragend', (e) => {
            if (e.target.classList.contains('subject-card')) {
                e.target.classList.remove('dragging');
            }
        });
        
        // 使用事件委托处理表格拖拽
        const timetable = document.getElementById('timetable');
        
        timetable.addEventListener('dragover', (e) => {
            const cell = e.target.closest('.cell');
            if (cell && !cell.classList.contains('occupied')) {
                e.preventDefault();
                cell.classList.add('drag-over');
            }
        });
        
        timetable.addEventListener('dragleave', (e) => {
            const cell = e.target.closest('.cell');
            if (cell) {
                cell.classList.remove('drag-over');
            }
        });
        
        timetable.addEventListener('drop', (e) => {
            const cell = e.target.closest('.cell');
            if (cell && !cell.classList.contains('occupied')) {
                e.preventDefault();
                cell.classList.remove('drag-over');
                
                if (this.draggedSubject) {
                    const day = cell.dataset.day;
                    const section = cell.dataset.section;
                    const period = cell.dataset.period;
                    this.addSubjectToCell(this.draggedSubject, day, section, period);
                }
            }
        });
        
        // 双击删除课程
        timetable.addEventListener('dblclick', (e) => {
            const cell = e.target.closest('.cell');
            if (cell && cell.classList.contains('occupied')) {
                this.removeSubjectFromCell(cell);
            }
        });
    }

    openSubjectModal(subject = null) {
        this.editingSubject = subject;
        const modal = document.getElementById('subjectModal');
        const nameInput = document.getElementById('subjectName');
        const teacherInput = document.getElementById('teacherName');
        const deleteBtn = document.getElementById('deleteSubjectBtn');
        
        if (subject) {
            nameInput.value = subject.name;
            teacherInput.value = subject.teacher;
            this.selectColorByValue(subject.color);
            deleteBtn.style.display = 'block';
        } else {
            nameInput.value = '';
            teacherInput.value = '';
            this.selectColorByValue('#FF6B6B');
            deleteBtn.style.display = 'none';
        }
        
        modal.style.display = 'block';
    }

    closeSubjectModal() {
        document.getElementById('subjectModal').style.display = 'none';
        this.editingSubject = null;
    }

    saveSubject(e) {
        e.preventDefault();
        
        const name = document.getElementById('subjectName').value.trim();
        const teacher = document.getElementById('teacherName').value.trim();
        const color = document.getElementById('customColor').value;
        
        if (!name) return;
        
        if (this.editingSubject) {
            this.editingSubject.name = name;
            this.editingSubject.teacher = teacher;
            this.editingSubject.color = color;
        } else {
            const subject = {
                id: Date.now().toString(),
                name,
                teacher,
                color
            };
            this.subjects.push(subject);
        }
        
        this.saveData();
        this.renderSubjects();
        this.renderTimetable();
        this.closeSubjectModal();
    }

    deleteSubject() {
        if (this.editingSubject) {
            this.deleteSubjectFromPool(this.editingSubject.id);
            this.closeSubjectModal();
        }
    }

    deleteSubjectFromPool(subjectId) {
        const subject = this.subjects.find(s => s.id === subjectId);
        if (subject) {
            // 从科目列表中删除
            this.subjects = this.subjects.filter(s => s.id !== subjectId);
            
            // 从课程表中移除该科目的所有实例
            Object.keys(this.timetable).forEach(key => {
                if (this.timetable[key] === subjectId) {
                    delete this.timetable[key];
                }
            });
            
            this.saveData();
            this.renderSubjects();
            this.renderTimetable();
        }
    }

    selectColor(e) {
        const color = e.target.dataset.color;
        document.getElementById('customColor').value = color;
        document.getElementById('customColorText').value = color;
        
        document.querySelectorAll('.color-option').forEach(opt => {
            opt.classList.remove('selected');
        });
        e.target.classList.add('selected');
    }

    selectColorByValue(color) {
        document.getElementById('customColor').value = color;
        document.getElementById('customColorText').value = color;
        
        document.querySelectorAll('.color-option').forEach(opt => {
            opt.classList.toggle('selected', opt.dataset.color === color);
        });
    }

    openTimeModal(e) {
        const timeText = e.target;
        const period = timeText.dataset.period;
        const modal = document.getElementById('timeModal');
        const timeInput = document.getElementById('timeRange');
        
        timeInput.value = timeText.textContent;
        timeInput.dataset.period = period;
        modal.style.display = 'block';
    }

    closeTimeModal() {
        document.getElementById('timeModal').style.display = 'none';
    }

    openTutorialModal() {
        document.getElementById('tutorialModal').style.display = 'flex';
    }

    closeTutorialModal() {
        document.getElementById('tutorialModal').style.display = 'none';
    }

    // 设置相关方法
    loadSettings() {
        const savedSettings = localStorage.getItem('timetableSettings');
        if (savedSettings) {
            this.settings = { ...this.settings, ...JSON.parse(savedSettings) };
        }
    }

    saveSettings() {
        localStorage.setItem('timetableSettings', JSON.stringify(this.settings));
    }

    applySettings() {
        // 应用晚上课时显示设置 - 直接通过ID查找并隐藏整个控制行
        const eveningControlLines = document.querySelectorAll('.period-control-line');
        
        eveningControlLines.forEach(controlLine => {
            const span = controlLine.querySelector('span');
            if (span && span.textContent.trim() === '晚上课时') {
                // 隐藏整个控制行（包括文本和按钮）
                controlLine.style.display = this.settings.showEvening ? 'flex' : 'none';
                controlLine.style.visibility = this.settings.showEvening ? 'visible' : 'hidden';
            }
        });

        // 应用周六、周日显示设置
        const saturdayCol = document.getElementById('saturdayCol');
        const sundayCol = document.getElementById('sundayCol');
        
        if (saturdayCol) {
            saturdayCol.style.display = this.settings.showSaturday ? 'table-cell' : 'none';
        }
        if (sundayCol) {
            sundayCol.style.display = this.settings.showSunday ? 'table-cell' : 'none';
        }

        // 更新课程表中的周末列 - 重新渲染后应用设置
        setTimeout(() => {
            const weekendCols = document.querySelectorAll('.weekend-col');
            weekendCols.forEach(col => {
                if (col.dataset.day === '6') {
                    col.style.display = this.settings.showSaturday ? 'table-cell' : 'none';
                } else if (col.dataset.day === '7') {
                    col.style.display = this.settings.showSunday ? 'table-cell' : 'none';
                }
            });
        }, 0);

        // 应用时间显示设置
        setTimeout(() => {
            const timeDisplays = document.querySelectorAll('.time-display');
            timeDisplays.forEach(display => {
                display.style.display = this.settings.showPeriodTime ? 'block' : 'none';
            });
        }, 0);

        this.renderTimetable();
    }

    openSettingsModal() {
        const modal = document.getElementById('settingsModal');
        const showEveningCheckbox = document.getElementById('showEvening');
        const showSaturdayCheckbox = document.getElementById('showSaturday');
        const showSundayCheckbox = document.getElementById('showSunday');
        const showPeriodTimeCheckbox = document.getElementById('showPeriodTime');

        showEveningCheckbox.checked = this.settings.showEvening;
        showSaturdayCheckbox.checked = this.settings.showSaturday;
        showSundayCheckbox.checked = this.settings.showSunday;
        showPeriodTimeCheckbox.checked = this.settings.showPeriodTime;

        modal.style.display = 'flex';
    }

    closeSettingsModal() {
        document.getElementById('settingsModal').style.display = 'none';
    }

    // 备份数据相关方法
    toggleBackupMenu(e) {
        e.stopPropagation();
        const menu = document.getElementById('backupMenu');
        const isVisible = menu.classList.contains('show');
        
        // 关闭所有其他下拉菜单
        this.closeAllDropdowns();
        
        if (!isVisible) {
            menu.classList.add('show');
            this.positionDropdown(menu, e.target);
        }
    }

    closeBackupMenu(e) {
        const menu = document.getElementById('backupMenu');
        const button = document.getElementById('backupBtn');
        
        if (!menu.contains(e.target) && !button.contains(e.target)) {
            menu.classList.remove('show');
            document.removeEventListener('click', this.closeBackupMenu.bind(this));
        }
    }

    // 导出数据
    exportData() {
        try {
            const data = {
                timetable: this.timetable,
                subjects: this.subjects,
                periods: this.periods,
                settings: this.settings,
                exportTime: new Date().toISOString(),
                version: '1.0'
            };
            
            const dataStr = JSON.stringify(data, null, 2);
            const dataBlob = new Blob([dataStr], { type: 'application/json' });
            
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `课程表备份_${new Date().toLocaleDateString().replace(/\//g, '-')}.json`;
            
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            this.showNotification('数据导出成功！', 'success');
            this.closeBackupMenu({ target: null });
        } catch (error) {
            console.error('导出数据失败:', error);
            this.showNotification('导出数据失败，请重试', 'error');
        }
    }

    // 导入数据
    importData() {
        document.getElementById('importFileInput').click();
        this.closeBackupMenu({ target: null });
    }

    // 处理文件导入
    handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;

        console.log('开始导入文件:', file.name, '大小:', file.size, 'bytes');

        if (!file.name.endsWith('.json')) {
            this.showNotification('请选择JSON格式的备份文件', 'error');
            return;
        }

        if (file.size === 0) {
            this.showNotification('文件为空，请选择有效的备份文件', 'error');
            return;
        }

        if (file.size > 10 * 1024 * 1024) { // 10MB限制
            this.showNotification('文件过大，请选择小于10MB的备份文件', 'error');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                console.log('文件读取成功，文件内容长度:', e.target.result.length);
                console.log('文件内容前100字符:', e.target.result.substring(0, 100));
                
                // 检查文件内容是否为空
                if (!e.target.result || e.target.result.trim() === '') {
                    this.showNotification('文件内容为空，请选择有效的备份文件', 'error');
                    return;
                }
                
                console.log('开始解析JSON...');
                const data = JSON.parse(e.target.result);
                console.log('JSON解析成功，数据类型:', typeof data);
                console.log('数据内容:', data);
                
                // 验证数据格式
                if (!this.validateImportData(data)) {
                    this.showNotification('备份文件格式不正确，请检查文件内容', 'error');
                    return;
                }
                
                // 确认导入
                if (confirm('导入数据将覆盖当前课表，是否继续？')) {
                    console.log('用户确认导入，开始加载数据...');
                    this.loadImportedData(data);
                    this.showNotification('数据导入成功！', 'success');
                } else {
                    console.log('用户取消导入');
                }
            } catch (error) {
                console.error('导入数据失败:', error);
                console.error('错误详情:', {
                    name: error.name,
                    message: error.message,
                    stack: error.stack
                });
                
                let errorMessage = '文件解析失败';
                if (error instanceof SyntaxError) {
                    errorMessage = `JSON格式错误: ${error.message}`;
                    console.error('JSON解析错误位置:', error.message);
                } else if (error.message) {
                    errorMessage = `导入失败: ${error.message}`;
                }
                this.showNotification(errorMessage, 'error');
            }
        };
        
        reader.onerror = (error) => {
            console.error('文件读取失败:', error);
            this.showNotification('文件读取失败，请重试', 'error');
        };
        
        reader.readAsText(file, 'UTF-8');
        // 清空文件输入，允许重复选择同一文件
        event.target.value = '';
    }

    // 验证导入数据格式
    validateImportData(data) {
        try {
            console.log('开始验证导入数据:', data);
            
            // 基本结构检查
            if (!data || typeof data !== 'object') {
                console.error('数据格式错误: 不是有效的对象');
                return false;
            }
            
            // 检查必要字段
            if (!Array.isArray(data.subjects)) {
                console.error('数据格式错误: subjects 不是数组，实际类型:', typeof data.subjects);
                return false;
            }
            
            if (typeof data.timetable !== 'object') {
                console.error('数据格式错误: timetable 不是对象，实际类型:', typeof data.timetable);
                return false;
            }
            
            if (typeof data.periods !== 'object') {
                console.error('数据格式错误: periods 不是对象，实际类型:', typeof data.periods);
                return false;
            }
            
            // 检查periods结构 - 更宽松的验证
            if (data.periods.morning && !Array.isArray(data.periods.morning)) {
                console.error('数据格式错误: periods.morning 不是数组');
                return false;
            }
            
            if (data.periods.afternoon && !Array.isArray(data.periods.afternoon)) {
                console.error('数据格式错误: periods.afternoon 不是数组');
                return false;
            }
            
            if (data.periods.evening && !Array.isArray(data.periods.evening)) {
                console.error('数据格式错误: periods.evening 不是数组');
                return false;
            }
            
            // 检查科目数据格式 - 更宽松的验证
            for (let i = 0; i < data.subjects.length; i++) {
                const subject = data.subjects[i];
                if (!subject || typeof subject !== 'object') {
                    console.error(`数据格式错误: 科目[${i}]不是对象:`, subject);
                    return false;
                }
                if (!subject.id && !subject.name) {
                    console.error(`数据格式错误: 科目[${i}]缺少必要字段:`, subject);
                    return false;
                }
            }
            
            console.log('数据验证通过，包含字段:', Object.keys(data));
            return true;
        } catch (error) {
            console.error('验证数据时出错:', error);
            return false;
        }
    }

    // 加载导入的数据
    loadImportedData(data) {
        try {
            console.log('开始加载导入数据...');
            
            // 恢复课表数据 - 确保是对象
            this.timetable = (data.timetable && typeof data.timetable === 'object') ? data.timetable : {};
            console.log('课表数据加载:', Object.keys(this.timetable).length, '个时间段');
            
            // 恢复科目数据 - 确保是数组并验证完整性
            this.subjects = Array.isArray(data.subjects) ? data.subjects : [];
            this.subjects = this.subjects.filter(subject => {
                if (!subject || typeof subject !== 'object') {
                    console.warn('过滤掉无效的科目数据:', subject);
                    return false;
                }
                if (!subject.id) {
                    subject.id = 'subject_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
                    console.log('为科目生成新ID:', subject.name, subject.id);
                }
                return true;
            });
            console.log('科目数据加载:', this.subjects.length, '个科目');
            
            // 恢复时间段数据 - 提供默认值
            this.periods = {
                morning: Array.isArray(data.periods?.morning) ? data.periods.morning : [
                    { name: '第1节', time: '08:00-08:40' },
                    { name: '第2节', time: '08:50-09:30' },
                    { name: '第3节', time: '10:00-10:40' },
                    { name: '第4节', time: '10:50-11:30' }
                ],
                afternoon: Array.isArray(data.periods?.afternoon) ? data.periods.afternoon : [
                    { name: '第1节', time: '14:00-14:40' },
                    { name: '第2节', time: '14:50-15:30' },
                    { name: '第3节', time: '15:40-16:20' }
                ],
                evening: Array.isArray(data.periods?.evening) ? data.periods.evening : [
                    { name: '第1节', time: '19:00-19:40' },
                    { name: '第2节', time: '19:50-20:30' }
                ]
            };
            console.log('时间段数据加载完成');
            
            // 恢复设置数据 - 提供默认值
            if (data.settings && typeof data.settings === 'object') {
                this.settings = { ...this.settings, ...data.settings };
                console.log('设置数据加载:', this.settings);
                // 应用设置
                this.applySettings();
            }
            
            // 重新渲染界面
            this.renderTimetable();
            this.renderSubjects();
            
            // 保存到本地存储
            this.saveData();
            localStorage.setItem('timetableSettings', JSON.stringify(this.settings));
            
            console.log('数据导入成功，界面已更新');
        } catch (error) {
            console.error('加载导入数据时出错:', error);
            this.showNotification('导入数据时发生错误', 'error');
        }
    }

    // 显示通知
    showNotification(message, type = 'info') {
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 12px 20px;
            border-radius: 6px;
            color: white;
            font-weight: 500;
            z-index: 10000;
            animation: slideInRight 0.3s ease-out;
            max-width: 300px;
            word-wrap: break-word;
        `;
        
        // 根据类型设置颜色
        switch (type) {
            case 'success':
                notification.style.backgroundColor = '#28a745';
                break;
            case 'error':
                notification.style.backgroundColor = '#dc3545';
                break;
            case 'warning':
                notification.style.backgroundColor = '#ffc107';
                notification.style.color = '#333';
                break;
            default:
                notification.style.backgroundColor = '#007bff';
        }
        
        notification.textContent = message;
        document.body.appendChild(notification);
        
        // 3秒后自动移除
        setTimeout(() => {
            if (notification.parentNode) {
                notification.style.animation = 'slideOutRight 0.3s ease-in';
                setTimeout(() => {
                    if (notification.parentNode) {
                        notification.parentNode.removeChild(notification);
                    }
                }, 300);
            }
        }, 3000);
    }

    saveSettings(e) {
        e.preventDefault();
        
        const showEveningCheckbox = document.getElementById('showEvening');
        const showSaturdayCheckbox = document.getElementById('showSaturday');
        const showSundayCheckbox = document.getElementById('showSunday');
        const showPeriodTimeCheckbox = document.getElementById('showPeriodTime');

        this.settings.showEvening = showEveningCheckbox.checked;
        this.settings.showSaturday = showSaturdayCheckbox.checked;
        this.settings.showSunday = showSundayCheckbox.checked;
        this.settings.showPeriodTime = showPeriodTimeCheckbox.checked;

        // 保存设置到本地存储
        localStorage.setItem('timetableSettings', JSON.stringify(this.settings));
        
        // 应用设置并重新渲染
        this.applySettings();
        this.closeSettingsModal();
    }

    // 立即开始创建课程表功能
    startCreatingTimetable() {
        // 关闭教程弹窗
        this.closeTutorialModal();
        
        // 如果在手机端，确保显示科目池
        if (window.innerWidth <= 768) {
            const subjectPool = document.querySelector('.subject-pool');
            if (subjectPool) {
                subjectPool.style.display = 'block';
            }
        }
        
        // 滚动到课程表顶部，确保用户看到操作区域
        const timetableContainer = document.querySelector('.timetable-container');
        if (timetableContainer) {
            timetableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
        // 如果没有科目，提示用户添加
        if (this.subjects.length === 0) {
            // 显示一个简短提示
            const hint = document.createElement('div');
            hint.innerHTML = `
                <div style="position: fixed; top: 20px; left: 50%; transform: translateX(-50%); 
                background: #4CAF50; color: white; padding: 15px 25px; border-radius: 8px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 10000; 
                animation: fadeInOut 3s ease-in-out;">
                    请点击左侧「+ 科目」按钮开始添加科目
                </div>
                <style>
                @keyframes fadeInOut {
                    0% { opacity: 0; top: 0; }
                    10% { opacity: 1; top: 20px; }
                    90% { opacity: 1; top: 20px; }
                    100% { opacity: 0; top: 0; }
                }
                </style>
            `;
            document.body.appendChild(hint);
            
            // 3秒后自动移除提示
            setTimeout(() => {
                if (hint.parentNode) {
                    hint.parentNode.removeChild(hint);
                }
            }, 3000);
        }
        
        // 如果有科目，但科目池在手机端被隐藏，提示用户如何操作
        if (this.subjects.length > 0 && window.innerWidth <= 768) {
            // 检查科目池是否可见
            const subjectPool = document.querySelector('.subject-pool');
            // 使用getComputedStyle来准确判断元素是否可见
            const computedStyle = window.getComputedStyle(subjectPool);
            if (subjectPool && (subjectPool.style.display === 'none' || computedStyle.display === 'none')) {
                // 显示一个简短提示
                const hint = document.createElement('div');
                hint.innerHTML = `
                    <div style="position: fixed; top: 20px; left: 50%; transform: translateX(-50%); 
                    background: #2196F3; color: white; padding: 15px 25px; border-radius: 8px; 
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 10000; 
                    animation: fadeInOut 3s ease-in-out;">
                        请从上方科目池中拖拽科目到课程表中
                    </div>
                    <style>
                    @keyframes fadeInOut {
                        0% { opacity: 0; top: 0; }
                        10% { opacity: 1; top: 20px; }
                        90% { opacity: 1; top: 20px; }
                        100% { opacity: 0; top: 0; }
                    }
                    </style>
                `;
                document.body.appendChild(hint);
                
                // 3秒后自动移除提示
                setTimeout(() => {
                    if (hint.parentNode) {
                        hint.parentNode.removeChild(hint);
                    }
                }, 3000);
            }
        }
    }

    saveTime(e) {
        e.preventDefault();
        
        const timeInput = document.getElementById('timeRange');
        const period = timeInput.dataset.period;
        const newTime = timeInput.value.trim();
        
        if (!newTime) return;
        
        document.querySelector(`[data-period="${period}"]`).textContent = newTime;
        this.saveData();
        this.closeTimeModal();
    }

    addSubjectToCell(subjectId, day, section, period) {
        const key = `${day}-${section}-${period}`;
        this.timetable[key] = subjectId;
        this.saveData();
        this.renderTimetable();
    }

    removeSubjectFromCell(cell) {
        if (cell.classList.contains('occupied')) {
            const day = cell.dataset.day;
            const section = cell.dataset.section;
            const period = cell.dataset.period;
            const key = `${day}-${section}-${period}`;
            delete this.timetable[key];
            this.saveData();
            this.renderTimetable();
        }
    }

    renderSubjects() {
        const pool = document.getElementById('subjectPool');
        pool.innerHTML = '';
        
        this.subjects.forEach(subject => {
            const card = document.createElement('div');
            card.className = 'subject-card';
            card.draggable = true;
            card.dataset.subjectId = subject.id;
            card.style.borderLeft = `4px solid ${subject.color}`;
            
            const teacherHtml = subject.teacher ? `<div class="teacher-name">${subject.teacher}</div>` : '';
            const subjectStyle = !subject.teacher ? 'style="line-height: 40px;"' : '';
            
            card.innerHTML = `
                <div class="subject-info">
                    <div class="subject-name" ${subjectStyle}>${subject.name}</div>
                    ${teacherHtml}
                </div>
                <div class="subject-actions">
                    <button class="btn-text edit-btn" title="编辑" data-action="edit">编辑</button>
                    <button class="btn-text delete-btn" title="删除" data-action="delete">删除</button>
                </div>
            `;
            
            // 编辑按钮事件
            card.querySelector('.edit-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                this.openSubjectModal(subject);
            });
            
            // 删除按钮事件
            card.querySelector('.delete-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                this.deleteSubjectFromPool(subject.id);
            });
            
            pool.appendChild(card);
        });
    }

    addPeriod(section) {
        const periods = this.periods[section];
        let defaultTime = '15:00-15:40';
        switch(section) {
            case 'morning':
                defaultTime = '09:00-09:40';
                break;
            case 'afternoon':
                defaultTime = '15:00-15:40';
                break;
            case 'evening':
                defaultTime = '19:00-19:40';
                break;
        }
        const newPeriod = {
            name: `第${periods.length + 1}节`,
            time: defaultTime
        };
        periods.push(newPeriod);
        this.saveData();
        this.renderTimetable();
    }

    removePeriod(section) {
        if (this.periods[section].length <= 1) {
            alert('至少需要保留一节课！');
            return;
        }
        
        let sectionName = '下午';
        switch(section) {
            case 'morning':
                sectionName = '上午';
                break;
            case 'afternoon':
                sectionName = '下午';
                break;
            case 'evening':
                sectionName = '晚上';
                break;
        }
        
        if (confirm(`确定要删除${sectionName}的最后一节课吗？`)) {
            this.periods[section].pop();
            
            // 清理对应的课程表数据
            const keysToDelete = [];
            for (let key in this.timetable) {
                if (key.includes(`-${section}-`)) {
                    const parts = key.split('-');
                    const periodIndex = parseInt(parts[2]);
                    if (periodIndex >= this.periods[section].length) {
                        keysToDelete.push(key);
                    }
                }
            }
            
            keysToDelete.forEach(key => {
                delete this.timetable[key];
            });
            
            this.saveData();
            this.renderTimetable();
        }
    }

    renderTimetable() {
        const tbody = document.getElementById('timetableBody');
        tbody.innerHTML = '';
        
        // 渲染上午
        if (this.periods.morning.length > 0) {
            this.periods.morning.forEach((period, index) => {
                const row = this.createPeriodRow('morning', index, period);
                tbody.appendChild(row);
            });
        }
        
        // 渲染下午
        if (this.periods.afternoon.length > 0) {
            this.periods.afternoon.forEach((period, index) => {
                const row = this.createPeriodRow('afternoon', index, period);
                tbody.appendChild(row);
            });
        }

        // 渲染晚上
        if (this.settings.showEvening && this.periods.evening.length > 0) {
            this.periods.evening.forEach((period, index) => {
                const row = this.createPeriodRow('evening', index, period);
                tbody.appendChild(row);
            });
        }
    }

    // 手机端选择科目功能
    showMobileSubjectSelector(day, section, period) {
        const cellKey = `${day}-${section}-${period}`;
        
        // 创建弹窗
        const modal = document.createElement('div');
        modal.className = 'mobile-subject-modal';
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 2000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;
        
        const content = document.createElement('div');
        content.style.cssText = `
            background: white;
            border-radius: 12px;
            padding: 0;
            max-width: 320px;
            width: 90%;
            max-height: 75vh;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        `;
        
        // 头部区域
        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px 20px 15px;
            border-bottom: 1px solid #eee;
        `;
        
        const title = document.createElement('h3');
        title.textContent = '选择科目';
        title.style.cssText = 'margin: 0; font-size: 18px; color: #333; font-weight: 600;';
        
        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = `
            background: none;
            border: none;
            font-size: 24px;
            cursor: pointer;
            color: #999;
            width: 30px;
            height: 30px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            transition: all 0.2s;
        `;
        closeBtn.onmouseover = () => closeBtn.style.background = '#f5f5f5';
        closeBtn.onmouseout = () => closeBtn.style.background = 'none';
        
        header.appendChild(title);
        header.appendChild(closeBtn);
        
        // 科目列表区域
        const listContainer = document.createElement('div');
        listContainer.style.cssText = 'max-height: 50vh; overflow-y: auto; padding: 15px 20px;';
        
        const list = document.createElement('div');
        list.style.cssText = 'display: flex; flex-direction: column; gap: 8px;';
        
        if (this.subjects.length === 0) {
            const emptyMessage = document.createElement('div');
            emptyMessage.style.cssText = `
                text-align: center;
                padding: 40px 20px;
                color: #999;
                font-size: 14px;
            `;
            emptyMessage.innerHTML = `
                <div style="font-size: 48px; margin-bottom: 10px;">📚</div>
                <div>暂无科目，请先添加科目</div>
                <button onclick="document.getElementById('addSubjectBtn').click(); this.closest('.mobile-subject-modal').remove();" 
                        style="margin-top: 10px; padding: 8px 16px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer;">
                    添加科目
                </button>
            `;
            listContainer.appendChild(emptyMessage);
        } else {
            this.subjects.forEach(subject => {
                const item = document.createElement('div');
                item.style.cssText = `
                    padding: 12px 15px;
                    border: 1px solid #eee;
                    border-radius: 8px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    transition: all 0.2s;
                    background: white;
                `;
                item.onmouseover = () => {
                    item.style.background = '#f8f9fa';
                    item.style.borderColor = '#007bff';
                };
                item.onmouseout = () => {
                    item.style.background = 'white';
                    item.style.borderColor = '#eee';
                };
                
                const colorBox = document.createElement('div');
                colorBox.style.cssText = `
                    width: 24px;
                    height: 24px;
                    border-radius: 50%;
                    background: ${subject.color};
                    flex-shrink: 0;
                `;
                
                const textContainer = document.createElement('div');
                textContainer.style.cssText = 'flex: 1;';
                
                const subjectName = document.createElement('div');
                subjectName.textContent = subject.name;
                subjectName.style.cssText = 'font-weight: 600; color: #333; margin-bottom: 2px;';
                
                const teacherName = document.createElement('div');
                teacherName.textContent = subject.teacher || '暂无老师';
                teacherName.style.cssText = 'font-size: 12px; color: #666;';
                
                textContainer.appendChild(subjectName);
                textContainer.appendChild(teacherName);
                
                item.appendChild(colorBox);
                item.appendChild(textContainer);
                
                item.addEventListener('click', () => {
                    this.addSubjectToCell(subject.id, day, section, period);
                    this.closeMobileSubjectModal(modal);
                });
                
                list.appendChild(item);
            });
            listContainer.appendChild(list);
        }
        
        // 底部按钮区域
        const footer = document.createElement('div');
        footer.style.cssText = `
            padding: 15px 20px 20px;
            border-top: 1px solid #eee;
            display: flex;
            gap: 10px;
        `;
        
        const addSubjectBtn = document.createElement('button');
        addSubjectBtn.textContent = '添加新科目';
        addSubjectBtn.style.cssText = `
            flex: 1;
            padding: 10px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s;
        `;
        addSubjectBtn.onmouseover = () => addSubjectBtn.style.background = '#0056b3';
        addSubjectBtn.onmouseout = () => addSubjectBtn.style.background = '#007bff';
        addSubjectBtn.addEventListener('click', () => {
            this.closeMobileSubjectModal(modal);
            setTimeout(() => this.openSubjectModal(), 300);
        });
        
        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = '取消';
        cancelBtn.style.cssText = `
            flex: 1;
            padding: 10px;
            background: #6c757d;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s;
        `;
        cancelBtn.onmouseover = () => cancelBtn.style.background = '#545b62';
        cancelBtn.onmouseout = () => cancelBtn.style.background = '#6c757d';
        cancelBtn.addEventListener('click', () => {
            this.closeMobileSubjectModal(modal);
        });
        
        footer.appendChild(addSubjectBtn);
        footer.appendChild(cancelBtn);
        
        // 组装弹窗
        content.appendChild(header);
        content.appendChild(listContainer);
        content.appendChild(footer);
        modal.appendChild(content);
        
        // CSS动画已在styles.css中定义，无需动态添加
        
        // 多种关闭方式
        const closeModal = () => this.closeMobileSubjectModal(modal);
        
        // 1. 点击关闭按钮
        closeBtn.addEventListener('click', closeModal);
        
        // 2. 点击背景区域
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                closeModal();
            }
        });
        
        // 3. 按ESC键关闭
        const handleEscape = (e) => {
            if (e.key === 'Escape') {
                closeModal();
                document.removeEventListener('keydown', handleEscape);
            }
        };
        document.addEventListener('keydown', handleEscape);
        
        // 4. 点击取消按钮
        cancelBtn.addEventListener('click', closeModal);
        
        // 防止滚动穿透，同时避免页面晃动
        const scrollBarWidth = window.innerWidth - document.documentElement.clientWidth;
        document.body.style.overflow = 'hidden';
        if (scrollBarWidth > 0) {
            document.body.style.paddingRight = scrollBarWidth + 'px';
        }
        
        // 显示弹窗
        document.body.appendChild(modal);
        
        // 保存引用以便关闭
        this.currentMobileModal = modal;
    }
    
    // 关闭手机端选择科目弹窗
    closeMobileSubjectModal(modal) {
        if (!modal) return;
        
        // 直接关闭弹窗，无动画效果
        if (modal.parentNode) {
            document.body.removeChild(modal);
        }
        document.body.style.overflow = '';
        document.body.style.paddingRight = '';
        
        this.currentMobileModal = null;
    }

    createPeriodRow(section, periodIndex, period) {
        const row = document.createElement('tr');
        
        // 时间列 - 只在第一节创建，使用rowSpan合并单元格
        if (periodIndex === 0) {
            const timeCell = document.createElement('td');
            timeCell.className = 'time-cell';
            let timeText = '';
            switch(section) {
                case 'morning':
                    timeText = '<div class="vertical-text"><span>上</span><span>午</span></div>';
                    break;
                case 'afternoon':
                    timeText = '<div class="vertical-text"><span>下</span><span>午</span></div>';
                    break;
                case 'evening':
                    timeText = '<div class="vertical-text"><span>晚</span><span>上</span></div>';
                    break;
            }
            timeCell.innerHTML = timeText;
            timeCell.rowSpan = this.periods[section].length;
            row.appendChild(timeCell);
        }
        
        // 课时列
        const periodCell = document.createElement('td');
        periodCell.className = 'period-cell';
        const timeDisplayStyle = this.settings.showPeriodTime ? 'display: block;' : 'display: none;';
        periodCell.innerHTML = `
                        <div class="period-name" data-section="${section}" data-period="${periodIndex}" style="cursor: pointer; font-weight: bold; color: #000;">
                            ${period.name}
                        </div>
                        <div class="time-display" data-section="${section}" data-period="${periodIndex}" style="cursor: pointer; font-size: 12px; color: #666; ${timeDisplayStyle}">
                            ${period.time}
                        </div>
                    `;
        
        // 添加课时名称和时间段点击事件
        periodCell.querySelector('.period-name').addEventListener('click', (e) => {
            this.openTimeModal(e, section, periodIndex);
        });
        periodCell.querySelector('.time-display').addEventListener('click', (e) => {
            this.openTimeModal(e, section, periodIndex);
        });
        
        row.appendChild(periodCell);
        
        // 周一到周日的格子
        const days = [1, 2, 3, 4, 5];
        if (this.settings.showSaturday) days.push(6);
        if (this.settings.showSunday) days.push(7);

        for (let day of days) {
            const cell = document.createElement('td');
            cell.className = 'cell';
            if (day >= 6) {
                cell.classList.add('weekend-col');
            }
            cell.dataset.day = day;
            cell.dataset.section = section;
            cell.dataset.period = periodIndex;
            
            const key = `${day}-${section}-${periodIndex}`;
            const subjectId = this.timetable[key];
            
            if (subjectId) {
                const subject = this.subjects.find(s => s.id === subjectId);
                if (subject) {
                    cell.classList.add('occupied');
                    const content = document.createElement('div');
                    content.className = 'cell-content';
                    content.style.backgroundColor = subject.color;
                    const teacherHtml = subject.teacher ? `<div class="teacher-name">${subject.teacher}</div>` : '';
                    const subjectStyle = !subject.teacher ? 'style="margin-bottom: 0;"' : '';
                    content.innerHTML = `
                        <div class="subject-name" ${subjectStyle}>${subject.name}</div>
                        ${teacherHtml}
                        <button class="delete-cell-btn" title="删除课程">×</button>
                    `;
                    cell.appendChild(content);
                    
                    // 添加删除按钮事件
                    content.querySelector('.delete-cell-btn').addEventListener('click', (e) => {
                        e.stopPropagation();
                        this.removeSubjectFromCell(cell);
                    });
                }
            } else {
                // 所有设备默认显示+号
                cell.classList.add('empty-cell');
                cell.style.cssText = 'position: relative; cursor: pointer;';
                
                // 使用CSS伪元素显示+号，确保默认显示
                const plusIndicator = document.createElement('div');
                plusIndicator.className = 'plus-indicator';
                plusIndicator.textContent = '+';
                plusIndicator.style.cssText = 'font-size: 24px; color: #ccc; display: flex; align-items: center; justify-content: center; width: 100%; height: 100%;';
                cell.appendChild(plusIndicator);
            }
            
            // 添加双击删除课程事件
            cell.addEventListener('dblclick', () => {
                if (cell.classList.contains('occupied')) {
                    this.removeSubjectFromCell(cell);
                }
            });
            
            // 添加点击选择
            cell.addEventListener('click', () => {
                this.editingCell = cell;
                document.querySelectorAll('.cell').forEach(c => c.classList.remove('selected'));
                cell.classList.add('selected');
                
                // 所有设备点击选择科目
                if (!cell.classList.contains('occupied')) {
                    this.showMobileSubjectSelector(day, section, periodIndex);
                }
            });
            
            row.appendChild(cell);
        }
        
        return row;
    }

    openTimeModal(e, section, periodIndex) {
        this.editingPeriod = { section, periodIndex };
        const modal = document.getElementById('timeModal');
        const nameInput = document.getElementById('periodName');
        const startHourSelect = document.getElementById('startHour');
        const startMinuteSelect = document.getElementById('startMinute');
        const endHourSelect = document.getElementById('endHour');
        const endMinuteSelect = document.getElementById('endMinute');
        
        const period = this.periods[section][periodIndex];
        nameInput.value = period.name;
        
        // 解析现有时间
        const [startTime, endTime] = period.time.split('-');
        const [startH, startM] = startTime.split(':');
        const [endH, endM] = endTime.split(':');
        
        startHourSelect.value = startH;
        startMinuteSelect.value = startM;
        endHourSelect.value = endH;
        endMinuteSelect.value = endM;
        
        modal.style.display = 'block';
    }

    savePeriodTime(e) {
        e.preventDefault();
        
        if (!this.editingPeriod) return;
        
        const { section, periodIndex } = this.editingPeriod;
        const nameInput = document.getElementById('periodName');
        const startHourSelect = document.getElementById('startHour');
        const startMinuteSelect = document.getElementById('startMinute');
        const endHourSelect = document.getElementById('endHour');
        const endMinuteSelect = document.getElementById('endMinute');
        
        const newName = nameInput.value.trim();
        const startHour = startHourSelect.value;
        const startMinute = startMinuteSelect.value;
        const endHour = endHourSelect.value;
        const endMinute = endMinuteSelect.value;
        
        if (!newName || !startHour || !startMinute || !endHour || !endMinute) return;
        
        const newTime = `${startHour}:${startMinute}-${endHour}:${endMinute}`;
        this.periods[section][periodIndex].time = newTime;
        this.periods[section][periodIndex].name = newName;
        this.saveData();
        this.renderTimetable();
        this.closeTimeModal();
    }

    resetTimetable() {
        if (confirm('确定要重置整个课程表吗？这将清空课程表内容但保留科目')) {
            // 只重置课程表内容，保留科目池
            this.timetable = {};
            this.periods = {
                morning: [
                { name: '第1节', time: '08:00-08:40' },
                { name: '第2节', time: '08:50-09:30' },
                { name: '第3节', time: '10:00-10:40' },
                { name: '第4节', time: '10:50-11:30' }
            ],
            afternoon: [
                { name: '第1节', time: '14:00-14:40' },
                { name: '第2节', time: '14:50-15:30' },
                { name: '第3节', time: '15:40-16:20' }
            ]
            };
            
            // 重置课程表标题
            const defaultTitle = '我的课程表';
            document.getElementById('timetableTitle').value = defaultTitle;
            localStorage.setItem('timetableTitle', defaultTitle);
            
            this.saveData();
            this.renderTimetable();
        }
    }

    toggleExportDropdown(e) {
        e.stopPropagation();
        const dropdown = document.getElementById('exportMenu');
        const isVisible = dropdown.classList.contains('show');
        
        // 关闭所有其他下拉菜单
        this.closeAllDropdowns();
        
        // 切换当前下拉菜单
        if (!isVisible) {
            dropdown.classList.add('show');
            this.positionDropdown(dropdown, e.target);
        }
    }

    closeAllDropdowns() {
        // 关闭导出菜单
        const exportMenu = document.getElementById('exportMenu');
        if (exportMenu) {
            exportMenu.classList.remove('show');
        }
        
        // 关闭备份菜单
        const backupMenu = document.getElementById('backupMenu');
        if (backupMenu) {
            backupMenu.classList.remove('show');
        }
        
        // 关闭其他下拉菜单
        const dropdowns = document.querySelectorAll('.dropdown-content');
        dropdowns.forEach(dropdown => {
            dropdown.style.display = 'none';
        });
        
        // 隐藏移动端遮罩层
        this.hideMobileOverlay();
    }
    
    // 处理全局点击事件
    handleGlobalClick(e) {
        const exportMenu = document.getElementById('exportMenu');
        const backupMenu = document.getElementById('backupMenu');
        const exportBtn = document.getElementById('exportBtn');
        const backupBtn = document.getElementById('backupBtn');
        
        // 检查是否点击在下拉菜单或按钮上
        const isClickOnExportMenu = exportMenu && (exportMenu.contains(e.target) || exportBtn.contains(e.target));
        const isClickOnBackupMenu = backupMenu && (backupMenu.contains(e.target) || backupBtn.contains(e.target));
        
        // 如果点击在外部区域，隐藏所有下拉菜单
        if (!isClickOnExportMenu && !isClickOnBackupMenu) {
            this.closeAllDropdowns();
        }
    }
    
    // 智能定位下拉菜单
    positionDropdown(dropdown, button) {
        if (!dropdown || !button) return;
        
        const isMobile = window.innerWidth <= 768;
        
        if (isMobile) {
            // 手机端：固定定位，避免被遮挡
            this.positionMobileDropdown(dropdown, button);
        } else {
            // PC端：相对定位
            this.positionDesktopDropdown(dropdown, button);
        }
    }
    
    // PC端下拉菜单定位
    positionDesktopDropdown(dropdown, button) {
        const rect = button.getBoundingClientRect();
        const dropdownRect = dropdown.getBoundingClientRect();
        
        // 重置样式
        dropdown.style.position = 'absolute';
        dropdown.style.top = '100%';
        dropdown.style.left = '0';
        dropdown.style.right = 'auto';
        dropdown.style.bottom = 'auto';
        dropdown.style.transform = 'none';
        dropdown.style.zIndex = '9999';
        
        // 检查是否需要调整位置避免超出屏幕
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        
        if (rect.left + dropdownRect.width > viewportWidth) {
            dropdown.style.left = 'auto';
            dropdown.style.right = '0';
        }
        
        if (rect.bottom + dropdownRect.height > viewportHeight) {
            dropdown.style.top = 'auto';
            dropdown.style.bottom = '100%';
        }
    }
    
    // 手机端下拉菜单定位
    positionMobileDropdown(dropdown, button) {
        // 手机端使用固定定位，从底部弹出
        dropdown.style.position = 'fixed';
        dropdown.style.top = 'auto';
        dropdown.style.bottom = '20px';
        dropdown.style.left = '20px';
        dropdown.style.right = '20px';
        dropdown.style.width = 'auto';
        dropdown.style.transform = 'translateY(120%)';
        dropdown.style.zIndex = '999999';
        dropdown.style.transition = 'transform 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
        dropdown.style.maxHeight = '50vh';
        dropdown.style.overflowY = 'auto';
        
        // 创建或显示遮罩层
        this.createMobileOverlay();
    }
    
    // 创建移动端遮罩层
    createMobileOverlay() {
        let overlay = document.querySelector('.dropdown-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.className = 'dropdown-overlay';
            document.body.appendChild(overlay);
            
            // 点击遮罩层关闭菜单
            overlay.addEventListener('click', () => {
                this.closeAllDropdowns();
            });
        }
        overlay.classList.add('show');
    }
    
    // 隐藏移动端遮罩层
    hideMobileOverlay() {
        const overlay = document.querySelector('.dropdown-overlay');
        if (overlay) {
            overlay.classList.remove('show');
        }
    }


    // 处理窗口大小改变
    handleWindowResize() {
        // 检查是否有打开的下拉菜单，重新定位
        const exportMenu = document.getElementById('exportMenu');
        const backupMenu = document.getElementById('backupMenu');
        const exportBtn = document.getElementById('exportBtn');
        const backupBtn = document.getElementById('backupBtn');
        
        if (exportMenu && exportMenu.classList.contains('show') && exportBtn) {
            this.positionDropdown(exportMenu, exportBtn);
        }
        
        if (backupMenu && backupMenu.classList.contains('show') && backupBtn) {
            this.positionDropdown(backupMenu, backupBtn);
        }
        
        // 如果窗口变大，关闭手机端侧边栏
        if (window.innerWidth > 768) {
            this.closeMobileSidebar();
        }
    }
    
    // 手机端侧边栏相关方法
    toggleMobileSidebar() {
        const sidebar = document.getElementById('mobileSidebar');
        const overlay = document.getElementById('sidebarOverlay');
        const hamburgerBtn = document.getElementById('hamburgerBtn');
        
        if (sidebar.classList.contains('show')) {
            this.closeMobileSidebar();
        } else {
            this.openMobileSidebar();
        }
    }
    
    openMobileSidebar() {
        const sidebar = document.getElementById('mobileSidebar');
        const overlay = document.getElementById('sidebarOverlay');
        const hamburgerBtn = document.getElementById('hamburgerBtn');
        
        sidebar.classList.add('show');
        overlay.classList.add('show');
        hamburgerBtn.classList.add('active');
        
        // 防止背景滚动
        document.body.style.overflow = 'hidden';
    }
    
    closeMobileSidebar() {
        const sidebar = document.getElementById('mobileSidebar');
        const overlay = document.getElementById('sidebarOverlay');
        const hamburgerBtn = document.getElementById('hamburgerBtn');
        
        sidebar.classList.remove('show');
        overlay.classList.remove('show');
        hamburgerBtn.classList.remove('active');
        
        // 恢复背景滚动
        document.body.style.overflow = '';
    }
    
    // 处理侧边栏菜单按钮点击
    handleSidebarAction(action) {
        // 关闭侧边栏
        this.closeMobileSidebar();
        
        // 根据action执行相应功能
        switch (action) {
            case 'tutorial':
                this.openTutorialModal();
                break;
            case 'reset':
                this.resetTimetable();
                break;
            case 'saveImage':
                this.saveAsImage();
                break;
            case 'exportWord':
                this.exportToWord();
                break;
            case 'exportExcel':
                this.exportToExcel();
                break;
            case 'exportData':
                this.exportData();
                break;
            case 'importData':
                this.importData();
                break;
            case 'settings':
                this.openSettingsModal();
                break;
            default:
                console.warn('未知的侧边栏操作:', action);
        }
    }

    saveAsImage() {
        try {
            // 检测是否为移动端
            const isMobile = window.innerWidth <= 768;
            
            // 创建一个干净的容器用于截图 - 移动端和PC端统一尺寸
            const cleanContainer = document.createElement('div');
            cleanContainer.style.cssText = `
                position: absolute; 
                top: -9999px; 
                left: -9999px; 
                width: 1000px; 
                padding: 30px; 
                background: white;
                font-family: "Microsoft YaHei", Arial, sans-serif;
            `;

            // 克隆表格内容
            const originalTable = document.querySelector('.timetable-container');
            const tableClone = originalTable.cloneNode(true);

            // 移除不需要的元素（上午/下午控制按钮）
            const controls = tableClone.querySelectorAll('.section-controls');
            controls.forEach(control => control.remove());
            
            // 移除原有标题和标题输入区域
            const existingTitles = tableClone.querySelectorAll('h1, h2, h3, .title, .timetable-title-section, .table-title-input');
            existingTitles.forEach(title => title.remove());

            // 使用表格标题作为图片标题
            const titleInput = document.getElementById('tableTitle');
            
            // 创建标题
            const titleDiv = document.createElement('div');
            titleDiv.style.cssText = 'text-align: center; margin-bottom: 20px; padding: 0;';
            
            const mainTitle = document.createElement('h1');
            mainTitle.textContent = titleInput.value || '课程表';
            mainTitle.style.cssText = 'margin: 0; font-size: 28px; font-weight: bold; color: #333; font-family: "Microsoft YaHei", Arial, sans-serif;';
            
            titleDiv.appendChild(mainTitle);
            cleanContainer.appendChild(titleDiv);
            cleanContainer.appendChild(tableClone);
            
            // 统一表格样式，确保移动端和PC端效果一致
            const table = tableClone.querySelector('.timetable') || tableClone;
            table.style.cssText = `
                border-collapse: collapse;
                width: 100%;
                table-layout: fixed;
                border: 2px solid #333;
                font-size: 14px;
            `;
            
            // 统一单元格样式
            const cells = table.querySelectorAll('td, th');
            cells.forEach(cell => {
                const isHeader = cell.tagName.toLowerCase() === 'th' || cell.classList.contains('time-header') || cell.classList.contains('period-header');
                cell.style.cssText = `
                    border: 1px solid #333;
                    padding: 12px 8px;
                    text-align: center;
                    vertical-align: middle;
                    min-width: 80px;
                    min-height: 60px;
                    font-family: "Microsoft YaHei", Arial, sans-serif;
                    font-size: 14px;
                    ${isHeader ? 'background-color: #f8f9fa; font-weight: bold;' : ''}
                `;
            });
            
            // 根据设置隐藏周六和周日列
            const rows = table.querySelectorAll('tr');
            rows.forEach(row => {
                const cells = row.querySelectorAll('td, th');
                let cellIndex = 0;
                
                cells.forEach(cell => {
                    // 跳过前两列（时间和课时列）
                    if (cellIndex >= 2) {
                        const dayIndex = cellIndex - 2; // 0:周一, 1:周二, ..., 5:周六, 6:周日
                        
                        // 隐藏周六列（索引5）
                        if (dayIndex === 5 && !this.settings.showSaturday) {
                            cell.style.display = 'none';
                        }
                        
                        // 隐藏周日列（索引6）
                        if (dayIndex === 6 && !this.settings.showSunday) {
                            cell.style.display = 'none';
                        }
                    }
                    cellIndex++;
                });
            });

            // 处理上午下午区域背景色
            const timeSections = table.querySelectorAll('td[rowspan]');
            timeSections.forEach(section => {
                if (section.textContent.includes('上午')) {
                    section.style.backgroundColor = '#e6f3ff';
                    section.style.fontWeight = 'bold';
                } else if (section.textContent.includes('下午')) {
                    section.style.backgroundColor = '#fff2e6';
                    section.style.fontWeight = 'bold';
                }
            });
            
            document.body.appendChild(cleanContainer);
            
            html2canvas(cleanContainer, {
                backgroundColor: '#ffffff',
                scale: isMobile ? 3 : 2, // 移动端使用更高分辨率确保清晰度
                useCORS: true,
                allowTaint: true,
                width: 1000,
                height: cleanContainer.scrollHeight,
                windowWidth: 1000
            }).then(canvas => {
                const link = document.createElement('a');
                link.download = `${titleInput.value || '课程表'}.png`;
                link.href = canvas.toDataURL('image/png', 1.0);
                link.click();
                
                // 清理临时容器
                document.body.removeChild(cleanContainer);
            }).catch(error => {
                console.error('生成图片失败:', error);
                alert('生成图片失败，请重试');
                document.body.removeChild(cleanContainer);
            });
        } catch (error) {
            console.error('保存图片出错:', error);
            alert('保存图片出错，请重试');
        }
    }



    saveData() {
        const data = {
            subjects: this.subjects,
            timetable: this.timetable,
            periods: this.periods
        };
        localStorage.setItem('timetableData', JSON.stringify(data));
    }

    loadData() {
        const data = localStorage.getItem('timetableData');
        let hasValidSubjects = false;
        
        if (data) {
            const parsed = JSON.parse(data);
            this.timetable = parsed.timetable || {};
            this.periods = parsed.periods || {
                morning: [
                    { name: '第1节', time: '08:00-08:40' },
                    { name: '第2节', time: '08:50-09:30' },
                    { name: '第3节', time: '10:00-10:40' },
                    { name: '第4节', time: '10:50-11:30' }
                ],
                afternoon: [
                    { name: '第1节', time: '14:00-14:40' },
                    { name: '第2节', time: '14:50-15:30' },
                    { name: '第3节', time: '15:40-16:20' }
                ],
                evening: [
                    { name: '第1节', time: '19:00-19:40' },
                    { name: '第2节', time: '19:50-20:30' }
                ]
            };
            
            // 确保evening存在
            if (!this.periods.evening) {
                this.periods.evening = [
                    { name: '第1节', time: '19:00-19:40' },
                    { name: '第2节', time: '19:50-20:30' }
                ];
            }
            
            // 检查是否有有效的科目数据
            if (parsed.subjects && Array.isArray(parsed.subjects) && parsed.subjects.length > 0) {
                this.subjects = parsed.subjects;
                hasValidSubjects = true;
            }
        } else {
            // 如果没有数据，初始化所有时段
            this.periods = {
                morning: [
                    { name: '第1节', time: '08:00-08:40' },
                    { name: '第2节', time: '08:50-09:30' },
                    { name: '第3节', time: '10:00-10:40' },
                    { name: '第4节', time: '10:50-11:30' }
                ],
                afternoon: [
                    { name: '第1节', time: '14:00-14:40' },
                    { name: '第2节', time: '14:50-15:30' },
                    { name: '第3节', time: '15:40-16:20' }
                ],
                evening: [
                    { name: '第1节', time: '19:00-19:40' },
                    { name: '第2节', time: '19:50-20:30' }
                ]
            };
        }
        
        // 如果没有有效科目数据，使用默认科目
        if (!hasValidSubjects) {
            this.subjects = [
                { id: '1', name: '语文', teacher: '', color: '#FF0000' },
                { id: '2', name: '数学', teacher: '', color: '#0000FF' },
                { id: '3', name: '英语', teacher: '', color: '#FF69B4' },
                { id: '4', name: '美术', teacher: '', color: '#FFA500' },
                { id: '5', name: '健康与体育', teacher: '', color: '#800080' },
                { id: '6', name: '道德与法制', teacher: '', color: '#B8860B' },
                { id: '7', name: '音乐', teacher: '', color: '#FF4500' },
                { id: '8', name: '劳动', teacher: '', color: '#808080' },
                { id: '9', name: '班队会', teacher: '', color: '#663399' },
                { id: '10', name: '值日', teacher: '', color: '#006400' },
                { id: '11', name: '历史', teacher: '', color: '#000080' },
                { id: '12', name: '地理', teacher: '', color: '#4B0082' },
                { id: '13', name: '政治', teacher: '', color: '#8B0000' },
                { id: '14', name: '物理', teacher: '', color: '#A0522D' },
                { id: '15', name: '化学', teacher: '', color: '#2F4F4F' },
                { id: '16', name: '生物', teacher: '', color: '#FF6B6B' }
            ];
            // 立即保存默认科目到本地存储
            this.saveData();
        }
    }

    loadTimetableTitle() {
        const savedTitle = localStorage.getItem('timetableTitle');
        const titleInput = document.getElementById('timetableTitle');
        if (savedTitle) {
            titleInput.value = savedTitle;
        }
    }

    saveTimetableTitle(title) {
        localStorage.setItem('timetableTitle', title);
    }

    saveTableTitle(title) {
        localStorage.setItem('tableTitle', title);
    }

    loadTableTitle() {
        const savedTitle = localStorage.getItem('tableTitle');
        const titleInput = document.getElementById('tableTitle');
        if (savedTitle) {
            titleInput.value = savedTitle;
        }
    }

    initTimeSelectors() {
        // 生成小时选项 (6-22)
        const startHourSelect = document.getElementById('startHour');
        const endHourSelect = document.getElementById('endHour');
        
        for (let i = 6; i <= 22; i++) {
            const hour = i.toString().padStart(2, '0');
            startHourSelect.appendChild(new Option(hour, hour));
            endHourSelect.appendChild(new Option(hour, hour));
        }
        
        // 生成分钟选项 (00-55，间隔5分钟)
        const startMinuteSelect = document.getElementById('startMinute');
        const endMinuteSelect = document.getElementById('endMinute');
        
        for (let i = 0; i < 60; i += 5) {
            const minute = i.toString().padStart(2, '0');
            startMinuteSelect.appendChild(new Option(minute, minute));
            endMinuteSelect.appendChild(new Option(minute, minute));
        }
    }

    // Word导出功能 - 移动端PC端统一效果
    exportToWord() {
        const title = document.getElementById('tableTitle').value || '课程表';
        
        // 检测是否为移动端
        const isMobile = window.innerWidth <= 768;
        
        // 创建兼容Word的HTML格式（移动端PC端统一）
        let wordContent = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <meta charset="utf-8">
            <title>${title}</title>
            <!--[if gte mso 9]>
            <xml>
                <w:WordDocument>
                    <w:View>Print</w:View>
                    <w:Zoom>100</w:Zoom>
                    <w:DoNotOptimizeForBrowser/>
                </w:WordDocument>
            </xml>
            <![endif]-->
            <style>
                @page { margin: 2cm; }
                body { font-family: 'Microsoft YaHei', 'SimSun', Arial, sans-serif; margin: 20px; }
                .main-title { text-align: center; color: #333; margin-bottom: 25px; font-size: 26px; font-weight: bold; }
                table { border-collapse: collapse; width: 100%; margin: 0 auto; table-layout: fixed; border: 2px solid #333; }
                th, td { border: 1px solid #333; padding: 12px 8px; text-align: center; font-size: 14px; vertical-align: middle; min-height: 60px; }
                th { background-color: #f8f9fa; font-weight: bold; }
                .time-header { background-color: #e6f3ff; font-weight: bold; width: 80px; font-size: 14px; }
                .period-header { background-color: #f0f0f0; font-weight: bold; width: 100px; font-size: 14px; }
                .time-section { background-color: #e6f3ff; font-weight: bold; vertical-align: middle; }
                .time-section.pm { background-color: #fff2e6; }
                .subject { font-weight: bold; color: #333; font-size: 14px; }
                .teacher { font-size: 12px; color: #666; margin-top: 4px; display: block; }
                .period-time { font-size: 12px; color: #666; }
            </style>
        </head>
        <body>
            <div class="main-title">${title}</div>
            <table>
                <thead>
                    <tr>
                        <th class="time-header">时间</th>
                        <th class="period-header">课时</th>
                        ${(() => {
                            let headers = ['周一', '周二', '周三', '周四', '周五'];
                            if (this.settings.showSaturday) headers.push('周六');
                            if (this.settings.showSunday) headers.push('周日');
                            return headers.map(day => `<th>${day}</th>`).join('');
                        })()}
                    </tr>
                </thead>
                <tbody>`;

        // 构建表格内容
        
        // 添加上午部分
        if (this.periods.morning && this.periods.morning.length > 0) {
            this.periods.morning.forEach((period, periodIndex) => {
                wordContent += `<tr>`;
                
                // 时间列（只在第一节显示）
                if (periodIndex === 0) {
                    wordContent += `<td class="time-header" rowspan="${this.periods.morning.length}">上午</td>`;
                }
                
                // 节数和时间
                wordContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><small>${period.time}</small>` : ''}</td>`;
                
                // 每天的课程（根据设置动态显示）
                const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                for (let day = 1; day <= dayCount; day++) {
                    const key = `${day}-morning-${periodIndex}`;
                    const subjectId = this.timetable[key];
                    
                    if (subjectId) {
                        const subject = this.subjects.find(s => s.id === subjectId);
                        if (subject) {
                            wordContent += `<td>
                                <div class="subject">${subject.name}</div>
                                <div class="teacher">${subject.teacher || ''}</div>
                            </td>`;
                        } else {
                            wordContent += `<td></td>`;
                        }
                    } else {
                        wordContent += `<td></td>`;
                    }
                }
                
                wordContent += `</tr>`;
            });
        }
        
        // 添加下午部分
        if (this.periods.afternoon && this.periods.afternoon.length > 0) {
            this.periods.afternoon.forEach((period, periodIndex) => {
                wordContent += `<tr>`;
                
                // 时间列（只在第一节显示）
                if (periodIndex === 0) {
                    wordContent += `<td class="time-header" rowspan="${this.periods.afternoon.length}">下午</td>`;
                }
                
                // 节数和时间
                wordContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><small>${period.time}</small>` : ''}</td>`;
                
                // 每天的课程（根据设置动态显示）
                const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                for (let day = 1; day <= dayCount; day++) {
                    const key = `${day}-afternoon-${periodIndex}`;
                    const subjectId = this.timetable[key];
                    
                    if (subjectId) {
                        const subject = this.subjects.find(s => s.id === subjectId);
                        if (subject) {
                            wordContent += `<td>
                                <div class="subject">${subject.name}</div>
                                <div class="teacher">${subject.teacher || ''}</div>
                            </td>`;
                        } else {
                            wordContent += `<td></td>`;
                        }
                    } else {
                        wordContent += `<td></td>`;
                    }
                }
                
                wordContent += `</tr>`;
            });
        }

        
        // 添加晚上部分（如果显示）
        if (this.settings.showEvening && this.periods.evening && this.periods.evening.length > 0) {
            this.periods.evening.forEach((period, periodIndex) => {
                wordContent += `<tr>`;
                
                // 时间列（只在第一节显示）
                if (periodIndex === 0) {
                    wordContent += `<td class="time-header" rowspan="${this.periods.evening.length}">晚上</td>`;
                }
                
                // 节数和时间
                wordContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><small>${period.time}</small>` : ''}</td>`;
                
                // 每天的课程（根据设置动态显示）
                const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                for (let day = 1; day <= dayCount; day++) {
                    const key = `${day}-evening-${periodIndex}`;
                    const subjectId = this.timetable[key];
                    
                    if (subjectId) {
                        const subject = this.subjects.find(s => s.id === subjectId);
                        if (subject) {
                            wordContent += `<td>
                                <div class="subject">${subject.name}</div>
                                <div class="teacher">${subject.teacher || ''}</div>
                            </td>`;
                        } else {
                            wordContent += `<td></td>`;
                        }
                    } else {
                        wordContent += `<td></td>`;
                    }
                }
                
                wordContent += `</tr>`;
            });
        }

        wordContent += `</tbody></table></body></html>`;

        // 创建Blob并下载（使用正确的HTML格式）
        const blob = new Blob([wordContent], { type: 'application/msword;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${title}.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }
    
    // Excel导出功能 - 移动端PC端统一效果
    exportToExcel() {
        try {
            const title = document.getElementById('tableTitle').value || '课程表';
            
            // 检测是否为移动端
            const isMobile = window.innerWidth <= 768;
            
            // 创建兼容Excel的HTML格式（移动端PC端统一）
            let excelContent = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
            <head>
                <meta charset="utf-8">
                <title>${title}</title>
                <style>
                    body { 
                        font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif; 
                        margin: 20px; 
                        background: white;
                    }
                    .main-title { 
                        text-align: center; 
                        color: #333; 
                        margin-bottom: 25px; 
                        font-size: 26px; 
                        font-weight: bold; 
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                    }
                    table { 
                        border-collapse: collapse; 
                        width: 800px; 
                        margin: 0 auto; 
                        table-layout: fixed; 
                        border: 2px solid #333;
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                    }
                    th, td { 
                        border: 1px solid #333; 
                        padding: 12px 8px; 
                        text-align: center; 
                        font-size: 14px; 
                        vertical-align: middle; 
                        min-height: 60px;
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                        mso-style-parent:style0;
                        mso-rotate:0;
                        mso-font-charset:134;
                        white-space: normal;
                        word-wrap: break-word;
                    }
                    th { 
                        background-color: #f8f9fa; 
                        font-weight: bold; 
                    }
                    .time-header { 
                        background-color: #e6f3ff; 
                        font-weight: bold; 
                        width: 80px; 
                        font-size: 14px;
                    }
                    .period-header { 
                        background-color: #f0f0f0; 
                        font-weight: bold; 
                        width: 100px; 
                        font-size: 14px;
                    }
                    .time-section { 
                        background-color: #e6f3ff; 
                        font-weight: bold; 
                        vertical-align: middle;
                    }
                    .time-section.pm { 
                        background-color: #fff2e6; 
                    }
                    .time-section.evening { 
                        background-color: #f0f8ff; 
                    }
                    .subject { 
                        font-weight: bold; 
                        color: #333; 
                        font-size: 14px;
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                    }
                    .teacher { 
                            font-size: 12px; 
                            color: #666; 
                            font-family: "Microsoft YaHei", Arial, sans-serif;
                        }
                    .period-time { 
                        font-size: 12px; 
                        color: #666; 
                        font-family: "Microsoft YaHei", Arial, sans-serif;
                    }
                </style>
            </head>
            <body>
                <div class="main-title">${title}</div>
                <table>
                    <thead>
                        <tr>
                            <th class="time-header">时间</th>
                            <th class="period-header">课时</th>
                            ${(() => {
                                let headers = ['周一', '周二', '周三', '周四', '周五'];
                                if (this.settings.showSaturday) headers.push('周六');
                                if (this.settings.showSunday) headers.push('周日');
                                return headers.map(day => `<th>${day}</th>`).join('');
                            })()}
                        </tr>
                    </thead>
                    <tbody>`;
            
            // 上午课程
            for (let i = 0; i < this.periods.morning.length; i++) {
                const period = this.periods.morning[i];
                excelContent += '<tr>';
                
                if (i === 0) {
                    excelContent += `<td rowspan="${this.periods.morning.length}" class="time-section">上午</td>`;
                }
                
                excelContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><span class="period-time">${period.time}</span>` : ''}</td>`;
                
                // 每天的课程（根据设置动态显示）
                const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                for (let day = 1; day <= dayCount; day++) {
                    const key = `${day}-morning-${i}`;
                    const subjectId = this.timetable[key];
                    
                    if (subjectId) {
                        const subject = this.subjects.find(s => s.id === subjectId);
                        if (subject) {
                            excelContent += `<td><span class="subject">${subject.name}</span>`;
                            if (subject.teacher) {
                                excelContent += `<br><span class="teacher">${subject.teacher}</span>`;
                            }
                            excelContent += `</td>`;
                        } else {
                            excelContent += '<td></td>';
                        }
                    } else {
                        excelContent += '<td></td>';
                    }
                }
                
                excelContent += '</tr>';
            }
            
            // 下午课程
            for (let i = 0; i < this.periods.afternoon.length; i++) {
                const period = this.periods.afternoon[i];
                excelContent += '<tr>';
                
                if (i === 0) {
                    excelContent += `<td rowspan="${this.periods.afternoon.length}" class="time-section pm">下午</td>`;
                }
                
                excelContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><span class="period-time">${period.time}</span>` : ''}</td>`;
                
                // 每天的课程（根据设置动态显示）
                const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                for (let day = 1; day <= dayCount; day++) {
                    const key = `${day}-afternoon-${i}`;
                    const subjectId = this.timetable[key];
                    
                    if (subjectId) {
                        const subject = this.subjects.find(s => s.id === subjectId);
                        if (subject) {
                            excelContent += `<td><span class="subject">${subject.name}</span>`;
                            if (subject.teacher) {
                                excelContent += `<br><span class="teacher">${subject.teacher}</span>`;
                            }
                            excelContent += `</td>`;
                        } else {
                            excelContent += '<td></td>';
                        }
                    } else {
                        excelContent += '<td></td>';
                    }
                }
                
                excelContent += '</tr>';
            }
            
            
            // 晚上课程（如果显示）
            if (this.settings.showEvening && this.periods.evening && this.periods.evening.length > 0) {
                for (let i = 0; i < this.periods.evening.length; i++) {
                    const period = this.periods.evening[i];
                    excelContent += '<tr>';
                    
                    if (i === 0) {
                        excelContent += `<td rowspan="${this.periods.evening.length}" class="time-section evening">晚上</td>`;
                    }
                    
                    excelContent += `<td class="period-header">${period.name}${this.settings.showPeriodTime ? `<br><span class="period-time">${period.time}</span>` : ''}</td>`;
                    
                    // 每天的课程（根据设置动态显示）
                    const dayCount = 5 + (this.settings.showSaturday ? 1 : 0) + (this.settings.showSunday ? 1 : 0);
                    for (let day = 1; day <= dayCount; day++) {
                        const key = `${day}-evening-${i}`;
                        const subjectId = this.timetable[key];
                        
                        if (subjectId) {
                            const subject = this.subjects.find(s => s.id === subjectId);
                            if (subject) {
                                excelContent += `<td><span class="subject">${subject.name}</span>`;
                                if (subject.teacher) {
                                    excelContent += `<br><span class="teacher">${subject.teacher}</span>`;
                                }
                                excelContent += `</td>`;
                            } else {
                                excelContent += '<td></td>';
                            }
                        } else {
                            excelContent += '<td></td>';
                        }
                    }
                    
                    excelContent += '</tr>';
                }
            }
            
            excelContent += '</tbody></table></body></html>';
            
            // 创建Excel文件（使用兼容的HTML格式）
            const blob = new Blob([excelContent], { type: 'application/vnd.ms-excel;charset=utf-8' });
            const link = document.createElement('a');
            link.download = `${title}.xls`;
            link.href = URL.createObjectURL(blob);
            link.click();
            URL.revokeObjectURL(link.href);
            
        } catch (error) {
            console.error('导出Excel出错:', error);
            alert('导出Excel出错，请重试');
        }
    }
}

// 初始化应用
let app;
document.addEventListener('DOMContentLoaded', () => {
    app = new TimetableApp();
});