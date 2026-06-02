# 2026年进口货物物流跟踪表

这个仓库把 Excel 工作簿转换成一个可用 GitHub Pages 公开访问的静态页面。页面支持在浏览器内编辑单元格，并可导出编辑后的 Excel 兼容文件或当前工作表 CSV。

## 文件

- `index.html`：网页入口
- `data/workbook.js`：从 Excel 提取的工作簿数据
- `assets/2026年进口货物物流跟踪表-铁矿协作一部&新拓.xlsx`：原始 Excel 文件
- `scripts/build_logistics_repo.py`：重新生成站点的脚本

## 编辑和导出

- 点击任意单元格即可编辑。
- 修改会自动保存在当前浏览器的本地存储中。
- `导出 Excel` 会导出包含全部工作表的 Excel 兼容 `.xls` 文件。
- `导出当前 CSV` 会导出当前标签页工作表。
- 页面内修改不会自动写回 GitHub 仓库；需要导出后重新上传或由维护者更新源文件。

## 发布到 GitHub Pages

创建 GitHub 公开仓库后，在本目录执行：

```powershell
git remote add origin https://github.com/<your-user>/<repo-name>.git
git branch -M main
git push -u origin main
```

然后在 GitHub 仓库页面打开 `Settings -> Pages`，选择 `Deploy from a branch`，分支选 `main`，目录选 `/root`。
公开链接会类似：

```text
https://<your-user>.github.io/<repo-name>/
```

注意：公开仓库和 GitHub Pages 会让仓库里的表格数据被互联网上任何人访问。
