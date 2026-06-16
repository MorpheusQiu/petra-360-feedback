# Petra品牌中国区 360度反馈在线调查系统

## 快速启动

```bash
cd 360-feedback
pip install flask flask-cors werkzeug openpyxl
python app.py
```

访问 http://localhost:5000

## 登录信息

### 管理员账号
| 英文名 | 初始密码 | 角色 |
|--------|---------|------|
| Morpheus | petra2026 | HR管理员 |
| Mursal | petra2026 | 管理员 |
| Ali | petra2026 | 管理员 |
| Rita | petra2026 | 管理员 |

### 员工账号
所有员工初始密码均为 `petra2026`，首次登录后可修改。
共29名预置员工，英文名见Excel人员清单。

## 系统架构

- **后端**: Flask + SQLite (app.py)
- **前端**: 纯HTML/CSS/JS 单页应用 (static/index.html)
- **数据库**: SQLite (feedback.db，首次运行自动创建)

## 功能模块

### 员工端
1. 英文名+密码登录
2. 四维度评价表单：
   - 维度一：员工互评（5维度+开放题，可添加多位同事）
   - 维度二：对直属上级反馈（6维度+开放题）
   - 维度三：上级对下级反馈（6维度+开放题，可添加多位下属）
   - 维度四：匿名领导力反馈（8维度+开放题+组织氛围题）
3. 自动保存（3秒防抖+2分钟定时）
4. 提交后锁定，需HR开放修改权限

### 管理员端
1. 仪表盘：四维度提交统计、部门分布
2. 数据查看：按维度查看所有提交记录
3. 数据导出：Excel/CSV格式
4. 人员管理：查看提交状态、开放修改权限、重置密码
5. 系统设置：开放/关闭提交、周期名称

## API 文档

| 端点 | 方法 | 说明 | 权限 |
|------|------|------|------|
| /api/login | POST | 员工登录 | 公开 |
| /api/me | GET | 获取当前用户信息 | 登录 |
| /api/change_password | POST | 修改密码 | 登录 |
| /api/users | GET | 用户列表 | 登录 |
| /api/managers | GET | 管理层列表 | 登录 |
| /api/dim1 | GET/POST | 维度一CRUD | 登录 |
| /api/dim2 | GET/POST | 维度二CRUD | 登录 |
| /api/dim3 | GET/POST | 维度三CRUD | 登录 |
| /api/dim4/token | POST | 获取匿名token | 登录 |
| /api/dim4 | GET/POST | 维度四CRUD | 登录 |
| /api/autosave | POST | 批量自动保存 | 登录 |
| /api/status | GET | 提交状态 | 登录 |
| /api/admin/dashboard | GET | 仪表盘统计 | 管理员 |
| /api/admin/submissions/<dim> | GET | 提交记录 | 管理员 |
| /api/admin/users | GET/PUT | 用户管理 | 管理员 |
| /api/admin/settings | GET/PUT | 系统设置 | 管理员 |
| /api/admin/unlock/<uid> | POST | 开放修改 | 管理员 |
| /api/admin/export/<dim> | GET | 数据导出 | 管理员 |
