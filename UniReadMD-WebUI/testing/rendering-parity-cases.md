# 渲染展示对照样本

用同一份文档核对 Typora 与 UniReadMD。请保持相同窗口宽度、内容缩放和侧栏状态。

## 普通正文与行内格式

普通正文第一行
普通正文第二行
普通正文第三行

连续空格：甲    乙。

中文与 English 混排，数字 123456，**加粗**、*斜体*、~~删除线~~、`行内代码`和[链接](https://typora.io)。

## 列表与引用

- 第一项
- 第二项
  - 嵌套列表

3. 从三开始
4. 下一项

- [ ] 待完成事项
- [x] 已完成事项

> 引用第一行
> 引用第二行
>
> 引用中的第二个段落。

## 表格

| 左对齐 | 居中 | 右对齐 |
| :--- | :---: | ---: |
| 普通文字 | **加粗** | 123.45 |
| 较长的中文单元格用于观察折行与列宽 | `code` | 678.90 |

## 代码高亮

```javascript
const message = '正文展示';
console.log(message);
```

```powershell
Get-Process | Where-Object { $_.CPU -gt 5 }
```

```csharp
public record Note(string Title);
```

## 公式与脚注

行内公式 $x^2 + y^2 = z^2$，以及化学式 $\ce{H2O}$。

$$
\int_0^1 x^2\,dx = \frac{1}{3}
$$

脚注示例[^display]。

[^display]: 观察脚注字号、跳转和返回位置。

## 图表

```mermaid
flowchart LR
  A[中文节点] --> B[渲染展示]
```
