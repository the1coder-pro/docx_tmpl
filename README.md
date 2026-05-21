# docx_tmpl

A powerful, lightweight Dart and Flutter package that allows you to use Jinja2-style templating directly inside Microsoft Word (`.docx`) files. 🚀

Design your highly styled documents visually in Microsoft Word 🎨, insert familiar templating tags 🏷️, and dynamically populate them using Dart code 💻. **docx_tmpl** processes the file while preserving all of your original formatting, typography, and layout styling perfectly! 💯

---

## ⚡ Features

* **Variables & Properties 📝:** Simple `{{ variable }}` replacement with nested object evaluation (e.g., `{{ client.name }}`).
* **Conditionals 🔀:** Conditionally display, hide, or swap layout blocks using `{% if condition %}` and `{% endif %}`.
* **Loops / Iteration 🔄:** Render dynamic lists and repeating rows with `{% for item in items %}` and `{% endfor %}`.
* **Native Word Tables 📊:** Seamlessly loop through table rows to dynamically generate complex billing or reporting tables.
* **Dynamic Images & Hyperlinks 🖼️🔗:** Inject images dynamically and insert clickable URLs right into your template tags.
* **100% Format Retention 💎:** Your document fonts, sizes, text alignments, borders, headers, and footers remain completely untouched.
* **Pure Dart & Platform Independent 🌐:** Zero native platform dependencies. It is completely platform-agnostic and works anywhere Dart runs—fully supporting **iOS, Android, Windows, macOS, Linux, and the Web**.

---

## 🚀 Getting Started

### 1. Installation 📥

Add `docx_tmpl` to your `pubspec.yaml` file:

```yaml
dependencies:
  docx_tmpl: ^1.0.0
```

### 2. Prepare Your Word Template 📄

Open Microsoft Word and design your document exactly how you want it to look. Insert tags inline anywhere in your text or table cells. ⚙️

> **Visual Template Example in Word 🖼️:**
> 
> # Invoice: {{ invoiceNumber }}
> **Customer:** {{ customer.name }}
> 
> {% if highValueClient %}
> 🌟 *Priority Support Enabled*
> {% endif %}
> 
> **Items:**
> | Description | Total |
> | :--- | :--- |
> | `{% for item in items %}`{{ item.name }} | ${{ item.price }} `{% endfor %}` |
> 
> **Receipt Image:** {{ invoiceImage }}
> **Support Link:** {{ helpLink }}

---

## 🛠️ Usage

### Basic Example 💡

Here is how easily you can read a template file, parse it with data, and save the output. ✨

```dart
import 'dart:io';
import 'package:docx_tmpl/docx_tmpl.dart';

void main() async {
  // 1. Load your Microsoft Word template file bytes 📂
  final templateBytes = await File('template.docx').readAsBytes();

  // 2. Initialize the DocxTmpl parser ⚙️
  final docx = DocxTmpl(templateBytes);

  // 3. Define your structured dataset 📊
  final Map<String, dynamic> data = {
    'invoiceNumber': '#INV-2026-001',
    'customer': {
      'name': 'عبد الله محمد', // Full Arabic Support 🇸🇦✨
    },
    'highValueClient': true,
    'items': [
      {'name': 'Cloud Infrastructure Setup 🚀', 'price': 1200},
      {'name': 'Cross-Platform Mobile App Module 📱', 'price': 2500},
    ],
    // Image insertion payload 🖼️
    'invoiceImage': DocxImage.fromBytes(
      imageBytes: await File('signature.png').readAsBytes(),
      width: 150,
      height: 50,
    ),
    // Hyperlink insertion payload 🔗
    'helpLink': DocxLink(
      text: 'Visit Client Dashboard',
      url: '[https://example.com/dashboard](https://example.com/dashboard)',
    ),
  };

  // 4. Render the document with your data 🏗️
  final List<int> outputBytes = await docx.render(data);

  // 5. Save the generated .docx file 💾
  await File('generated_invoice.docx').writeAsBytes(outputBytes);
  print('Document generated successfully! 🎉🥳');
}
```

---

## 🔍 Advanced Syntax Guide

### Loops inside Tables 📊🔄
To populate a dynamic table, place the `{% for %}` tag in the first cell of the row you want to repeat, and the `{% endfor %}` tag at the end of the last cell in that same row. The engine automatically clones the row formatting for every item in the array. 🔥

### Conditionals 🛠️🔀
Conditional blocks can wrap text elements, table rows, or full sections. If the expression evaluates to `false` or `null`, everything enclosed between the `if` and `endif` tags is entirely stripped from the document rendering pipeline without leaving awkward empty vertical spaces. 🪄

---
## 📜 Credits & Inspiration

This package brings the powerful, industry-standard paradigm of `.docx` template manipulation into the Dart and Flutter ecosystem. 

* **Inspired by python-docx-template (`docxtpl`):** The conceptual framework, syntax rules, and layout design workflow of `docx_tmpl` are heavily inspired by the brilliant implementation of the popular Python package [docxtpl](https://github.com/elapouya/python-docx-template). It brings that same reliable Jinja2-style developer experience to Dart.
* **Attribution:** Please see the accompanying [CREDITS.md](CREDITS.md) file for more details about original package's info that was the base that were the start of building this package.  
