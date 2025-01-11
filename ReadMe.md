# README

## Last Versions

[win x86 v1.2.7](https://drive.google.com/drive/folders/1_CzBZpYRbT9xrK0BeX7R2gz-WcZuWGNm)

[win x86 v1.2.6](https://drive.google.com/drive/folders/1q1XsX6GDU9ZGJbYoo29hsUUcctb5l2dt)

[win x86 v1.2.5](https://drive.google.com/drive/folders/1sEfML7VFftfxHX5k8JWRD7VN8RXnc0dW)

[win x86 v1.2.4](https://drive.google.com/drive/folders/16R07amNl3zTPtK9Q3KcJj00oelNWgGoa)

[win x86 v1.2.3](https://drive.google.com/drive/folders/1L_iZZFv43d3aKCLD5kqQa7rrfKlSuxSl)

## Run comands:

build: `dotnet publish -c Release -r win-x86 --self-contained true /p:PublishSingleFile=true`

## Style and Page Settings Configuration

This document explains the configuration settings for style formatting (`[STYLE_NAME]` sections) and page settings (`{PAGE}` section) in documents for standardized formatting and layout requirements.

---

### Page Settings `{PAGE}`

The `{PAGE}` section configures general layout settings of the document. Options include page size, orientation, and margins.

#### Example:

```plaintext
{PAGE}
pageSize=A4
orientation=portrait
marginTop=2.5sm
marginBottom=1.5sm
marginRight=2sm
marginLeft=2sm
marginHeader=1.25sm
marginFooter=1.25sm
```

#### Parameters:

- **pageSize**: Defines the page size. Options: `A3`, `A4`, `A5`, `letter`.
- **orientation**: Page orientation. Options: `portrait`, `landscape`.
- **marginTop**, **marginBottom**, **marginRight**, **marginLeft**: Set the page margins (in cm).
- **marginHeader**, **marginFooter**: Margin space between header/footer and the page edge (in cm).

---

### Style Settings `[STYLE_NAME]`

Each `[STYLE_NAME]` section specifies settings for a particular style. These are used to control font size, color, alignment, spacing, and other text properties.

#### Example:

```plaintext
[ЕОМ: Назва таблиці]
name=ЕОМ: Назва таблиці
size=14
position=Center
lineSpacing=1.5
lineSpacingBefore=0
lineSpacingAfter=0
color=000000
fontType=Times New Roman
bold=true
italic=true
underline=false
capitalize=false
after="ЕОМ: Номер таблиці"
before="ЕОМ: Таблиця_центр", "ЕОМ: Таблиця_лів"
```

#### Parameters:

- **name**: Identifier for the style.
- **size**: Font size.
- **position**: Text alignment. Options: `Left`, `Center`, `Right`, `Both`.
- **lineSpacing**: Space between lines.
- **lineSpacingBefore**: Space before the paragraph.
- **lineSpacingAfter**: Space after the paragraph.
- **color**: Font color in hexadecimal format (e.g., `000000` for black).
- **fontType**: Font family (e.g., `Times New Roman`).
- **bold**: Enables bold text. Options: `true`, `false`.
- **italic**: Enables italic text. Options: `true`, `false`.
- **underline**: Enables underline. Options: `true`, `false`.
- **capitalize**: Capitalizes text. Options: `true`, `false`.
- **after**: Style(s) that should follow this style (comma-separated).
- **before**: Style(s) that should precede this style (comma-separated).

---

### Sample Configuration Overview

Below is an overview of frequently used styles:

- **ЕОМ: Назва таблиці**: A centered style for table titles, bold and italicized.
- **ЕОМ: Рисунок**: Centered style for figure captions, regular text.
- **ЕОМ: Автор**: Centered author name style, bold and capitalized.
- **ЕОМ: Основний**: Justified main body style without any special text attributes.

This configuration enables consistent formatting and clear document structure, ensuring that each section and element follows predefined styling and layout standards.
