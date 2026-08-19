"""
SolidWorks 参数编辑器 - 交互式终端工具

用法:
    python sw_cli.py                 # 交互式扫描零件并选择
    python sw_cli.py "C:\\path\\to\\part.SLDPRT"   # 直接指定零件

交互命令:
    变量=值            修改参数（支持中文标签，如: 边沿孔数量=3、总长=120）
    list               列出所有方程式和尺寸
    rollback           回退到基线快照（不写盘，需 save 保存）
    snapshot           保存当前状态为新基线
    save               保存修改到原文件
    save-as <名称>     另存为新 SLDPRT（不碰原件）
    step <名称>        导出 STEP
    export <名称>      另存 SLDPRT + 导出 STEP
    switch             切换/重新选择零件
    help               显示帮助
    exit / quit        退出

示例:
    > 边沿孔数量=3
    > 总长=120
    > export 加长款-3孔
"""
import sys
import os
import json
import time
import glob
import math
import win32com.client

swDocPART = 1
swOpenDocOptions_Silent = 1
SNAPSHOT_DIRNAME = ".sw_snapshots"


def _snapshot_path(sldprt_path):
    d = os.path.dirname(sldprt_path)
    name = os.path.splitext(os.path.basename(sldprt_path))[0]
    snap_dir = os.path.join(d, SNAPSHOT_DIRNAME)
    os.makedirs(snap_dir, exist_ok=True)
    return os.path.join(snap_dir, f"{name}_last.json")


class SWParamCLI:
    def __init__(self, sldprt_path):
        self.sldprt_path = sldprt_path
        try:
            self.sw = win32com.client.Dispatch("SldWorks.Application")
        except Exception as e:
            print("  ❌ 无法连接 SolidWorks。请先启动 SolidWorks 软件再运行本工具。")
            print(f"     错误: {e}")
            sys.exit(1)
        self.sw.Visible = True
        self.model = self._open_model()
        self.title = self.model.GetTitle
        self.eq_mgr = self.model.GetEquationMgr
        self.eq_names = self._get_equation_names()
        self.dim_map = self._build_dim_name_map()
        self.label_map = self._load_labels()
        self.dirty = False
        self._ensure_snapshot()

    # ---------- 基础 ----------
    def _open_model(self):
        model = self.sw.OpenDoc(self.sldprt_path, swDocPART)
        if model is None:
            try:
                model = self.sw.OpenDoc5(self.sldprt_path, swDocPART, swOpenDocOptions_Silent, "", 0)
            except Exception as e:
                print(f"ERROR: Cannot open {self.sldprt_path}: {e}")
                sys.exit(1)
        try:
            self.sw.ActivateDoc(self.sldprt_path)
        except Exception:
            pass
        return model

    def _get_equation_names(self):
        names = set()
        try:
            for i in range(self.eq_mgr.GetCount):
                eq = self.eq_mgr.Equation(i)
                if self.eq_mgr.GlobalVariable(i):
                    names.add(eq.split("=")[0].strip().strip('"').strip("'"))
        except Exception:
            pass
        return names

    def _build_dim_name_map(self):
        dim_map = {}
        feat = self.model.FirstFeature
        while feat is not None:
            self._collect_dim_names(feat, dim_map)
            sub = feat.GetFirstSubFeature
            while sub is not None:
                self._collect_dim_names(sub, dim_map)
                sub = sub.GetNextSubFeature
            feat = feat.GetNextFeature
        return dim_map

    def _collect_dim_names(self, feat, dim_map):
        try:
            disp_dim = feat.GetFirstDisplayDimension
        except Exception:
            return
        while disp_dim is not None:
            try:
                dim = disp_dim.GetDimension2(0, 0)
                if dim is not None:
                    full = dim.FullName
                    parts = full.split("@")
                    if len(parts) >= 2:
                        short = f"{parts[0]}@{parts[1]}"
                        if short not in dim_map:
                            dim_map[short] = full
                        if parts[0] not in dim_map:
                            dim_map[parts[0]] = full
            except Exception:
                try:
                    dim = disp_dim.GetDimension
                    if dim is not None:
                        full = dim.FullName
                        parts = full.split("@")
                        if len(parts) >= 2:
                            short = f"{parts[0]}@{parts[1]}"
                            if short not in dim_map:
                                dim_map[short] = full
                            if parts[0] not in dim_map:
                                dim_map[parts[0]] = full
                except Exception:
                    pass
            try:
                disp_dim = feat.GetNextDisplayDimension(disp_dim)
            except Exception:
                break

    def _load_labels(self):
        """Load Chinese labels from params_labeled.json next to the SLDPRT."""
        d = os.path.dirname(self.sldprt_path)
        p = os.path.join(d, "params_labeled.json")
        label_map = {}
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for eq_name, info in data.get("equations", {}).items():
                    label = info.get("label", "")
                    if label:
                        label_map[label] = eq_name
                for dim_name, info in data.get("dimensions", {}).items():
                    label = info.get("label", "")
                    if label:
                        label_map[label] = dim_name
            except Exception as e:
                print(f"(警告: params_labeled.json 解析失败: {e})")
        return label_map

    def _ensure_snapshot(self):
        snap_path = _snapshot_path(self.sldprt_path)
        if not os.path.exists(snap_path):
            self._save_snapshot()
            print("(已保存原始快照，rollback 可恢复)")
        else:
            print("(快照已存在)")

    def _save_snapshot(self):
        snap = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "sldprt_path": self.sldprt_path,
            "equations": {},
            "dimensions": {},
        }
        try:
            for i in range(self.eq_mgr.GetCount):
                snap["equations"][str(i)] = {
                    "equation": self.eq_mgr.Equation(i),
                    "value": self.eq_mgr.Value(i),
                    "is_global": bool(self.eq_mgr.GlobalVariable(i)),
                }
        except Exception:
            pass
        feat = self.model.FirstFeature
        while feat is not None:
            self._collect_dim_values(feat, snap["dimensions"])
            sub = feat.GetFirstSubFeature
            while sub is not None:
                self._collect_dim_values(sub, snap["dimensions"])
                sub = sub.GetNextSubFeature
            feat = feat.GetNextFeature
        path = _snapshot_path(self.sldprt_path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(snap, f, indent=2, ensure_ascii=False)
        return snap

    def _collect_dim_values(self, feat, dims):
        try:
            disp_dim = feat.GetFirstDisplayDimension
        except Exception:
            return
        while disp_dim is not None:
            try:
                dim = disp_dim.GetDimension2(0, 0)
                if dim is not None:
                    dims[dim.FullName] = round(dim.SystemValue * 1000, 6)
            except Exception:
                try:
                    dim = disp_dim.GetDimension
                    if dim is not None:
                        dims[dim.FullName] = round(dim.SystemValue * 1000, 6)
                except Exception:
                    pass
            try:
                disp_dim = feat.GetNextDisplayDimension(disp_dim)
            except Exception:
                break

    # ---------- 解析 ----------
    def _resolve(self, key):
        """Resolve user input to a real parameter name.
        Priority: exact equation var > label > exact dim > short dim > fuzzy label.
        """
        key = key.strip()
        if not key:
            return None
        if key in self.eq_names:
            return key
        if key in self.label_map:
            return self.label_map[key]
        if self.model.Parameter(key) is not None:
            return key
        if key in self.dim_map:
            return self.dim_map[key]
        # fuzzy: 标签包含匹配
        for label, target in self.label_map.items():
            if key in label or label in key:
                return target
        return None

    # ---------- 修改 ----------
    def modify(self, key, value):
        # 数值有限性校验（NaN/Inf/负尺寸都会破坏模型）
        if not math.isfinite(value) or value <= 0:
            return f"  ❌ 值必须为正的有限数字: {value}"
        target = self._resolve(key)
        if target is None:
            return f"  ❌ 找不到参数: {key}（用 list 查看可用参数）"

        # 方程式全局变量
        if target in self.eq_names:
            for i in range(self.eq_mgr.GetCount):
                eq = self.eq_mgr.Equation(i)
                if self.eq_mgr.GlobalVariable(i):
                    var = eq.split("=")[0].strip().strip('"').strip("'")
                    if var == target:
                        # 检测公式驱动变量（含运算符，或引用其他变量如 "A"="B"）
                        rhs = eq.split("=", 1)[1] if "=" in eq else ""
                        comment = ""
                        if "'" in eq:
                            comment = "'" + eq.split("'", 1)[1]
                        if any(op in rhs for op in ("+", "-", "*", "/", "(")) or '"' in rhs:
                            print(f"  ⚠️ {target} 是公式驱动变量 ({rhs.strip()})，直接改会破坏公式！")
                            return "  已取消。建议改它的驱动变量。"
                        old = self.eq_mgr.Value(i)
                        self.eq_mgr.Equation(i, f'"{target}" = {value}{comment}')
                        self.model.EditRebuild3
                        self.dirty = True
                        return f"  ✅ [方程式] {target}: {old} -> {value} {comment}"
            return f"  ❌ 找不到方程式: {target}"

        # 尺寸
        param = self.model.Parameter(target)
        if param is None and target in self.dim_map:
            # 短名映射: "L" -> "L@凸台-拉伸1@Part.Part"
            target = self.dim_map[target]
            param = self.model.Parameter(target)
        if param is None:
            return f"  ❌ 无法访问: {target}"
        old = param.SystemValue * 1000
        param.SystemValue = value / 1000.0
        self.model.EditRebuild3
        self.dirty = True
        return f"  ✅ {target}: {old:.2f} -> {value} mm"

    def show_equations(self):
        print(f"\n=== 方程式 ({self.eq_mgr.GetCount}) ===")
        for i in range(self.eq_mgr.GetCount):
            eq = self.eq_mgr.Equation(i)
            val = self.eq_mgr.Value(i)
            gv = " [全局变量]" if self.eq_mgr.GlobalVariable(i) else ""
            print(f"  {eq}  =>  {val}{gv}")

    def list_all(self):
        self.show_equations()
        print(f"\n=== 尺寸 ({len(self.dim_map)} 个短名) ===")
        print("  用短名或中文标签修改，如: 边沿孔数量=3")
        if self.label_map:
            print("\n=== 中文标签 ===")
            for label, target in sorted(self.label_map.items()):
                print(f"  {label} -> {target}")

    # ---------- 保存 ----------
    def save(self):
        ok = None
        try:
            self.model.Save3(0, None, None)
            ok = True
        except Exception:
            try:
                self.model.Save3(0)
                ok = True
            except Exception:
                try:
                    self.model.Save2(False)
                    ok = True
                except Exception:
                    try:
                        ok = self.model.Save
                    except Exception:
                        ok = False
        if ok:
            self.dirty = False
            print(f"  💾 已保存: {self.title}")
        else:
            print(f"  ❌ 保存失败: {self.title}（请检查 SolidWorks 状态）")

    def save_as(self, name):
        # 默认输出到零件目录的 output/ 子目录
        path = os.path.abspath(name)
        has_ext = os.path.splitext(path)[1]
        is_abs = os.path.isabs(name) or ":" in name
        if not has_ext:
            if is_abs and os.path.dirname(path):
                # 用户给了绝对路径目录（如 D:\out\名称）→ 用该目录
                base_dir = os.path.dirname(path)
            else:
                # 纯名称 → 零件目录的 output/
                base_dir = os.path.dirname(self.sldprt_path)
                path = os.path.join(base_dir, "output", name)
            os.makedirs(os.path.dirname(path), exist_ok=True)
        if not path.lower().endswith(".sldprt"):
            path += ".SLDPRT"
        out_dir = os.path.dirname(path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        try:
            ok = self.model.SaveAs(path)
        except Exception:
            try:
                ok = self.model.SaveAs4(path, 0, 0, None)
            except Exception as e:
                print(f"  ❌ SaveAs 错误: {e}")
                return
        if ok:
            # 另存成功后模型已指向新文件：更新状态，
            # 让后续 save/switch(CloseDoc)/快照/rollback 都针对新文件
            self.sldprt_path = path
            self.title = os.path.splitext(os.path.basename(path))[0]
            print(f"  💾 已另存: {path}")
        else:
            print(f"  ❌ 另存失败: {path}")

    def export_step(self, name):
        path = os.path.abspath(name)
        has_ext = os.path.splitext(path)[1]
        is_abs = os.path.isabs(name) or ":" in name
        if not has_ext:
            if is_abs and os.path.dirname(path):
                base_dir = os.path.dirname(path)
            else:
                base_dir = os.path.dirname(self.sldprt_path)
                path = os.path.join(base_dir, "output", name)
            os.makedirs(os.path.dirname(path), exist_ok=True)
        if not path.lower().endswith(".step"):
            path += ".STEP"
        out_dir = os.path.dirname(path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        try:
            ok = self.model.SaveAs(path)
        except Exception:
            try:
                ok = self.model.SaveAs4(path, 0, 1, None)
            except Exception as e:
                print(f"  ❌ STEP 导出错误: {e}")
                return
        if ok:
            print(f"  📦 已导出 STEP: {path}")
        else:
            print(f"  ❌ STEP 导出失败: {path}")

    # ---------- 回退 ----------
    def rollback(self):
        snap_path = _snapshot_path(self.sldprt_path)
        if not os.path.exists(snap_path):
            print("  ❌ 没有快照，无法回退")
            return
        try:
            with open(snap_path, "r", encoding="utf-8") as f:
                snap = json.load(f)
        except Exception:
            print("  ❌ 快照损坏，无法回退")
            return
        n_eq = n_dim = 0
        for idx_str, info in snap.get("equations", {}).items():
            try:
                idx = int(idx_str)
                cur = self.eq_mgr.Equation(idx)
                if cur != info["equation"]:
                    self.eq_mgr.Equation(idx, info["equation"])
                    n_eq += 1
            except Exception:
                pass
        for full, val in snap.get("dimensions", {}).items():
            try:
                p = self.model.Parameter(full)
                if p is not None and abs(p.SystemValue * 1000 - val) > 0.0001:
                    p.SystemValue = val / 1000.0
                    n_dim += 1
            except Exception:
                pass
        self.model.EditRebuild3
        self.dirty = True
        print(f"  ↩️  已回退 {n_eq} 个方程式, {n_dim} 个尺寸（用 save 写盘）")

    # ---------- 主循环 ----------
    def run(self):
        print("=" * 56)
        print(f"  SolidWorks 参数编辑器")
        print(f"  零件: {self.title}")
        print(f"  方程式: {self.eq_mgr.GetCount} 条 | 尺寸: {len(self.dim_map)} 个")
        print("  输入 help 查看命令, exit 退出")
        print("=" * 56)

        while True:
            try:
                line = input("\n> ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\n退出。")
                break
            if not line:
                continue

            line = line.strip()
            if not line:
                continue

            # 先按 "=" 切分：参数=值 是最常用的输入，允许等号两边有空格
            if "=" in line:
                key, _, val_str = line.partition("=")
                key = key.strip()
                val_str = val_str.strip()
                # 排除 save-as/export 等带 = 的命令（理论上没有，安全起见）
                if key.lower() in ("save-as", "export", "step"):
                    cmd = key.lower()
                    rest = val_str
                else:
                    try:
                        value = float(val_str)
                    except ValueError:
                        print(f"  ❌ 值必须是数字: {val_str}")
                        continue
                    msg = self.modify(key, value)
                    print(msg)
                    if "✅" in msg:
                        self.show_equations()
                    continue

            cmd = line.split()[0].lower()
            rest = line[len(cmd):].strip()

            if cmd in ("exit", "quit", "q"):
                if self.dirty:
                    try:
                        ans = input("  有未保存的修改，保存吗? [y/N] ").strip().lower()
                    except (EOFError, KeyboardInterrupt):
                        ans = "n"
                    if ans in ("y", "yes"):
                        self.save()
                print("退出。")
                break
            elif cmd == "help":
                print(__doc__)
            elif cmd == "list" or cmd == "ls":
                self.list_all()
            elif cmd == "rollback":
                self.rollback()
            elif cmd == "snapshot":
                self._save_snapshot()
                print("  📸 已保存当前状态为新基线")
            elif cmd == "save":
                self.save()
            elif cmd == "save-as":
                if rest:
                    self.save_as(rest)
                else:
                    print("  ❌ 用法: save-as <文件名>")
            elif cmd == "step":
                if rest:
                    self.export_step(rest)
                else:
                    print("  ❌ 用法: step <文件名>")
            elif cmd == "export":
                if rest:
                    # 统一计算输出路径：先剥离后缀，再复用 save_as 的目录逻辑
                    base = os.path.splitext(rest)[0]
                    # 计算目标目录：绝对路径目录 / 或零件 output/
                    if (os.path.isabs(rest) or ":" in rest) and os.path.dirname(os.path.abspath(rest)):
                        out_dir = os.path.dirname(os.path.abspath(rest))
                    else:
                        out_dir = os.path.join(os.path.dirname(self.sldprt_path), "output")
                    os.makedirs(out_dir, exist_ok=True)
                    p_sldprt = os.path.join(out_dir, base)
                    p_step = os.path.join(out_dir, base)
                    self.save_as(p_sldprt)
                    self.export_step(p_step)
                else:
                    print("  ❌ 用法: export <文件名>")
            elif cmd == "switch" or cmd == "cd":
                # 切换零件
                if self.dirty:
                    try:
                        ans = input("  有未保存的修改，保存吗? [y/N] ").strip().lower()
                    except (EOFError, KeyboardInterrupt):
                        ans = "n"
                    if ans in ("y", "yes"):
                        self.save()
                path = choose_part()
                if path:
                    # 关闭当前文档，避免累积
                    try:
                        self.sw.CloseDoc(self.title)
                    except Exception:
                        pass
                    self.__init__(path)
                    print(f"  已切换到: {self.title}")
            else:
                print(f"  ❓ 未知命令: {cmd}（输入 help 查看）")


def choose_part(search_root=None):
    """Scan for .SLDPRT files and let user pick one.

    search_root: 指定扫描目录；None 时用应用目录的上级（项目文件夹）。
    """
    app_dir = os.path.abspath(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if search_root is None:
        search_root = os.path.dirname(app_dir)  # 应用目录的上级 = 项目文件夹

    parts = []
    seen = set()
    for d in [search_root]:
        if not d or not os.path.exists(d):
            continue
        # 只扫 2 层：search_root 本身 + 一层子目录（防止递归扫全盘）
        glob_patterns = [
            os.path.join(d, "*.SLDPRT"),
            os.path.join(d, "*", "*.SLDPRT"),
        ]
        for pattern in glob_patterns:
            for p in glob.glob(pattern):
                name = os.path.basename(p)
                # 跳过 output 目录变体、临时锁文件、隐藏文件、应用自身目录
                if os.sep + "output" + os.sep in p:
                    continue
                if name.startswith("~$") or name.startswith("."):
                    continue
                if os.path.abspath(p).startswith(app_dir):
                    continue
                real = os.path.normcase(os.path.abspath(p))
                if real not in seen:
                    seen.add(real)
                    parts.append(os.path.abspath(p))

    if not parts:
        print("  ❌ 没找到 .SLDPRT 零件文件")
        print("  用法: python sw_cli.py \"C:\\path\\to\\part.SLDPRT\"")
        return None

    print("\n找到以下 SolidWorks 零件:")
    for i, p in enumerate(parts, 1):
        name = os.path.basename(p)
        d = os.path.dirname(p)
        print(f"  [{i}] {name}")
        print(f"      {d}")
    print(f"  [0] 退出")

    while True:
        try:
            choice = input("\n选择零件编号: ").strip()
        except (EOFError, KeyboardInterrupt):
            return None
        if not choice:
            continue
        if choice == "0":
            return None
        try:
            idx = int(choice)
        except ValueError:
            print("  ❌ 请输入编号")
            continue
        if 1 <= idx <= len(parts):
            return parts[idx - 1]
        print(f"  ❌ 编号超出范围 (1-{len(parts)})")


if __name__ == "__main__":
    if len(sys.argv) >= 2 and os.path.exists(sys.argv[1]):
        arg = sys.argv[1]
        if os.path.isdir(arg):
            # 参数是目录：扫描该目录
            path = choose_part(arg)
            if path is None:
                sys.exit(0)
        else:
            path = arg
        cli = SWParamCLI(path)
    else:
        path = choose_part()
        if path is None:
            sys.exit(0)
        cli = SWParamCLI(path)
    cli.run()
