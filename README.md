# OneNote Exporter

Thinking of moving your OneNote collection to another note-taking app such as Obsidian, Logseq, Org Mode and more? You're in the right place!

OneNote Exporter (in short, `one`) is a PowerShell program which is capable of exporting all your OneNote notes to any [Pandoc-supported plain text markup format](https://pandoc.org/MANUAL.html) using the OneNote Object Model and Pandoc. That is to say: markdown, org-mode and more!

## 2026 recovery update

This update keeps the original `one.ps1` exporter and adds a recovery companion at `tools\Invoke-OneNoteRecovery.ps1`. The recovery script is for exports where an older OneNote workflow produced `OneNote-Export-failed-pages.log` entries because OneNote COM publish failed, Windows path lengths were too long, long file names exceeded what OneNote could publish reliably, or HTML/PDF files were written into a recovery folder instead of the original export tree.

Recommended Windows setup in 2026:

* Windows 10 or 11
* Windows PowerShell 5.1 for OneNote COM automation
* OneNote desktop installed and opened before running exports
* Microsoft Word desktop and Pandoc for the main `one.ps1` markdown/org export path

Normal export:

```powershell
Copy-Item .\config_example.ps1 .\config.ps1
notepad .\config.ps1
.\one.ps1
```

Recovery after a failed PDF/HTML export log exists:

```powershell
.\tools\Invoke-OneNoteRecovery.ps1 -Step All
```

If your export folders are not under Downloads, pass the paths explicitly:

```powershell
.\tools\Invoke-OneNoteRecovery.ps1 `
  -Step All `
  -FailedLog "D:\OneNote\OneNote-Export-failed-pages.log" `
  -ExportRoot "D:\OneNote\OneNote-Export" `
  -RecoveryRoot "D:\OneNote\OneNote-Export-Recovered" `
  -ShortRoot "D:\OneNote\ONR2"
```

The recovery steps can also be run one at a time:

```powershell
.\tools\Invoke-OneNoteRecovery.ps1 -Step RetryV1
.\tools\Invoke-OneNoteRecovery.ps1 -Step RetryV2
.\tools\Invoke-OneNoteRecovery.ps1 -Step Merge
.\tools\Invoke-OneNoteRecovery.ps1 -Step FixV1Misclassified
.\tools\Invoke-OneNoteRecovery.ps1 -Step FixV1Flat
```

Recovered output is indexed in `OneNote-Export\_recovered\_recovered-index.csv`. Use `-Overwrite` only when you intentionally want to replace previous recovered copies.

### How the rescue process works

The recovery tool reads `OneNote-Export-failed-pages.log` and uses the saved OneNote page IDs to publish those failed pages again. This helps when the original export mostly worked but some pages failed because of long file names, deeply nested section paths, or OneNote COM path limits.

The `RetryV1` step first publishes each failed page into a temporary recovery folder and then copies the result back to the intended export location when possible. If the original location is still too long or problematic, it keeps the file in `OneNote-Export-Recovered` with shortened path components.

The `RetryV2` step retries the remaining failures through short temporary paths by creating a temporary `subst` drive such as `O:`. That gives OneNote a much shorter destination path, which is useful for notebooks with long page titles, long file names, or deeply nested section groups.

The `Merge` step copies successful rescue output into `OneNote-Export\_recovered` and writes `_recovered-index.csv`, so every rescued file can be traced back to the original failed target path. The two fix steps handle older V1 recovery layouts and flatten any remaining difficult HTML/PDF copies into short folders.

---
**Notable alternatives**

- [onenote-md-exporter](https://github.com/alxnbl/onenote-md-exporter)

---

### Table of Contents

[**Introduction**](#introduction)

[**Results**](#results)

[**Supported Markups**](#supported-markups)

[**Markup Packs**](#markup-packs)

[**Requirements**](#requirements)

[**Usage**](#usage)

[**Recommendations**](#recommendations)

[**Attribution**](#attribution)

---

## Introduction

`one` exports OneNote pages to Word using the OneNote Object Model, and then uses Pandoc to convert them to your markup format of choice. Then, `one` uses [**Markup Packs**](https://github.com/alopezrivera/one/tree/master/src/Conversion/Markup-Packs) to customize the result. Markup Packs are *functions specific to each markup format*, which contain search and replace queries executed at runtime against the text output by Pandoc to tailor it to your desires. If search and replace doesn't cut it, you can add a `postprocessing` scriptblock to increase your freedom. Markup Packs give you fine-grained control over of all elements of your notes, including

* Headers
* Metadata (eg: note creation date)
* Other markup elements such as horizontal lines, custom indentation and formatting, and whatever else you might be able to conjure up from the text in your notes

`one` currently ships Markup Packs for [Emacs Org Mode](https://github.com/alopezrivera/one/blob/master/src/Conversion/Markup-Packs/Org.psm1) (`OrgPack1`) and [markdown](https://github.com/alopezrivera/one/blob/master/src/Conversion/Markup-Packs/markdown.psm1) (`MarkDownPack1`).

### What is being exported?

`one` will export all your *local* OneNote notebooks, meaning that to export a notebook of yours, you will need to download it to OneNote >= 2016[*](#requirements) with the "Add Notebook" option.

### Customizing the output

As long as Pandoc supports your desired markup format, all `one` needs to shine is a Markup Pack to tailor the output to your tastes. [The section on Markup Packs](https://github.com/alopezrivera/one#adding-markup-packs) contains a step by step guide to write and use your own Markup Packs.

## Results

![OneNote test note along Org Mode and markdown exports](test/test.png)

You can see the actual test results in the [`test` directory](https://github.com/alopezrivera/one/tree/master/test) (as well as the Word file to which the test note was exported). I have attempted to identify all unsupported syntax, which you can see as you would in OneNote at the bottom of the [test Word file](https://github.com/alopezrivera/one/blob/master/test/test.docx), and the respective export (failure) in the [Org Mode](https://github.com/alopezrivera/one/blob/master/test/one-test.org) and [markdown](https://github.com/alopezrivera/one/blob/master/test/one-test.md) conversions.

As you can see in the image above, the Markup Packs shipping for Org Mode and markdown (`OrgPack1` and `MarkdownPack1` respectively) will give your notes:

* Note creation data (in the case of Org Mode in its timestamp format)
* Correctly rendered lists, numbered and unnumbered, as well as indented paragraphs
* And finally clean the output of export artifacts, excess newlines, etc

Some notes:

- *If you want markdown output compatible with VSCode and GitHub*, specify `markdown_github` in the [line 66](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1#L66) of your `config.ps1`

    ```
    $conversion = 'markdown_github-simple_tables-multiline_tables-grid_tables+pipe_tables'
    ```
* Formatting using different fonts and colors doesn't survive export, as could be expected
* Underscored text is annotated as such in markdown, but does not render correctly (at least in VSCode)
* Images resized within OneNote are rendered with size information when exporting to markdown. Be mindful of the markdown flavour you are using. Pandoc markdown (`markdown` in the [Pandoc call in your config.ps1](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1#L66)) image size notation will not render properly in GitHub or other GitHub-flavoured markdown renderers such as the VSCode markdown preview window.

## Supported Markups

With support is meant that `one` understands which file type you are trying to export your notes to: it will use this knowledge to appropriately name files and apply [default Markup Packs](#markup-packs) if `markupPack` is set to `''` in [line 74 of your config.ps1](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1#L74).

`one` supports all (as of June 2022) Pandoc supported markups, as follows (from the [Pandoc manual](https://pandoc.org/MANUAL.html)),

- Emacs Org Mode
  - `org`

- Markdown
  - `markdown_strict`

- CommonMark
  - `commonmark`
  - `commonmark_x`

- GitHub-Flavored Markdown
  - `gfm`
  - `markdown_github`

- Pandoc Markdown
  - `markdown`

- MultiMarkdown
  - `markdown_mmd`
        
- PHP Markdown Extra
  - `markdown_phpextra`

## Markup Packs

You can specify your Markup Pack of choice [line 74 of your config.ps1](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1#L74). `markupPack` may have three values, as follows:

### Configuration

#### `'<markup pack>'`

You Markup Pack of choice.

#### `''`

The default Markup Pack for your export format. `one` determines which Markup Pack to use by first [identifying the extension](https://github.com/alopezrivera/one/blob/7a6e7f9769eb8a05ca9e8f169699cd21fff55761/src/Conversion/Conversion-Markup.psm1#L3) of the file format you have specified in your [Pandoc call](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1#L66) (currently `.org` and `.md`), and then choosing the [default Markup Pack](https://github.com/alopezrivera/one/blob/7a6e7f9769eb8a05ca9e8f169699cd21fff55761/src/Conversion/Conversion-Markup.psm1#L94) for that format.

#### `'none'`

No post-processing will be applied.

### Adding Markup Packs

Markup Packs are *markup-format-specific* **functions** containing search and replace queries executed at runtime against a string containing the entire markup content. If search and replace doesn't cut it, you can add a `postprocessing` scriptblock to increase your freedom (check the scriptblock to "Remove over-indentation of list items" in [Markdown MarkdownPack1](https://github.com/alopezrivera/one/blob/master/src/Conversion/Markup-Packs/Markdown.psm1)).

A Markup Pack template is available in the [`templates` directory](https://github.com/alopezrivera/one/tree/master/templates). It's an annotated version of the [Emacs Org Mode **OrgPack1**](https://github.com/alopezrivera/one/blob/master/src/Conversion/Markup-Packs/Org.psm1) Markup Pack. If you're interested in exporting to a Markdown format, check the [Markdown MarkdownPack1](https://github.com/alopezrivera/one/blob/master/src/Conversion/Markup-Packs/Markdown.psm1) Markup Pack for inspiration.

To add a Markup Pack, follow these steps:

1. Write your Markup Pack in the file containing the Markup Packs of your markup format of choice (`Org.psm1` or `Markdown.psm1` in `src/Conversion/Markup-Packs`). 
2. Set `markupPack` in your [config.ps1](https://github.com/alopezrivera/one/blob/6ec09267553cec5848c02fa2f20531185b2b2289/config_example.ps1) to the name of your markup pack. That is, the name of the **function** you have written.

## Requirements

* Windows >= 10

* Windows Powershell 5.x and above, or [Powershell Core 6.x up to 7.0](https://github.com/PowerShell/PowerShell)

* Microsoft OneNote
  > \>= 2016 (Desktop version, **NOT the Windows Store version**)
  * Download: FREE - https://www.onenote.com/Download

* Microsoft Word
  > \>= 2016 (Desktop version, **NOT the Windows Store version**)
  * Download: Office 365 Trial - https://www.microsoft.com/en-us/microsoft-365/try

* [PanDoc >= 2.2.3.2](https://pandoc.org/installing.html)

  * TIP: You may also use [Chocolatey](https://chocolatey.org/docs/installation#install-with-powershellexe) to install Pandoc on Windows, this will also set the right path (environment) statements. (https://chocolatey.org/packages/pandoc)

## Usage

1. Clone this repository
1. Start the OneNote desktop application
1. Rename `config_example.ps1` to `config.ps1` and configure the available options to your liking.
1. Open a PowerShell terminal at the directory containing the script and run it.
      * `.\one.ps1`
1. Sit back and wait until the process completes. To stop the process at any time, press Ctrl+C.
* **While running the conversion OneNote will be unusable**, as the Object Model might be interrupted if OneNote is used through the conversion process.

### Options

All of the following are configured from `config.ps1` (assuming you have renamed `config example.ps1` to that).

* Create a **folder structure** for your Notebooks and Sections
  * Process pages that are in sections at the **Notebook, Section Group and all Nested Section Group levels**
* Choose between converting a **specific notebook** or **all notebooks**
* Choose between creating **subfolders for subpages** (e.g. `Page\Subpage.md`) or **appending prefixes** (e.g. `Page_Subpage.md`)
* Specify a value between `32` and `255` as the maximum length of markdown file names, and their folder names (only when using subfolders for subpages (e.g. `Page\Subpage.md`)). A lower value can help avoid hitting [file and folder name limits of `255` bytes on file systems](https://en.wikipedia.org/wiki/Comparison_of_file_systems#Limits). A higher value preserves a longer page title. If using page prefixes (e.g. `Page_Subpage.md`), it is recommended to use a value of `100` or greater.
* Choose between putting all media (images, attachments) in a central `/media` folder for each notebook, or in a separate `/media` folder in each folder of the hierarchy
  * Symbols in media file names removed for link compatibility
  * Updates media references in the resulting `.md` files, generating **relative** references to the media files within the markdown document
* Choose between **discarding or keeping intermediate Word files**. Intermediate Word files are stored in a central notebook folder.
* Choose between converting from existing `.docx` (90% faster) and creating new ones - useful if just want to test differences in the various processing options without generating new `.docx` each time
* Choose between naming `.docx` files using page ID and last modified epoch date e.g. `{somelongid}-1234567890.docx` or hierarchy e.g. `<sectiongroup>-<section>-<page>.docx`
* **Input the Pandoc call, including conversion format and any extensions**, defaulting to Pandoc markdown format which strips most HTML from tables and using pipe tables. [See more details on these options here](https://pandoc.org/MANUAL.html#options). Default configurations are provided in `config example.ps1`. The following formats are accepted, among others:
  * org (Emacs Org Mode)
  * markdown (Pandoc’s markdown)
  * commonmark (CommonMark markdown)
  * gfm (GitHub-Flavored markdown), or the deprecated and less accurate markdown_github; use markdown_github only if you need extensions not supported in gfm.
  * markdown_mmd (Multimarkdown)
  * markdown_phpextra (PHP markdown Extra)
  * markdown_strict (original unextended markdown)
* Choose whether to use a **default Markup Pack, a specific one, or none** if you want to remove all post-processing (useful for debugging purposes).
* Choose whether to include a page timestamp and separator at top of the page.
* Choose whether to remove double spaces between numbered and unnumbered lists, excess whitespace after list markers, non-breaking spaces from blank lines, and `>` after bullet lists, created by Pandoc
* Choose whether to remove `\` escape symbol that are created when converting with Pandoc
* Choose whether to use Line Feed (LF) or Carriage Return + Line Feed (CRLF) for new lines
* Choose whether to include a `.pdf` export alongside the `.md` file. `.md` does not preserve `InkDrawing` (i.e. overlayed drawings, highlights, pen marks) absolute positions within a page, but a `.pdf` export is a complete page snapshot that preserves `InkDrawing` absolute positions within a page.

## Recommendations

1. You may want to consider using VS Code and its embedded Powershell terminal, as this allows you to edit and run your configuration and check conversion results. To make things easier, consider setting `$notesdestpath` in `config.ps1` to a `notes` directory in the project while adjusting the settings to your preference.
1. If you aren't actively editing your pages in OneNote, it is highly recommended that you don't delete the intermediate Word docs, as their generation takes a large part of runtime. They are stored in their own folder, out of the way. You can then quickly re-run the script with different parameters until you find what you like.
1. If you happen to collapse paragraphs in OneNote, consider installing [Onetastic](https://getonetastic.com/download) and the [attached macro](https://github.com/alopezrivera/one/blob/master/Expand%20All%20Paragraphs%20in%20Notebook.xml), which will automatically expand any collapsed paragraphs in the notebook. They won't be exported otherwise.
   * To install the macro, click the New Macro Button within the Onetastic Toolbar and then select File -> Import and select the .xml macro included in the release.
   * Run the macro for each Notebook that is open
1. Unlock all password-protected sections before continuing, the Object Model will not have access to them otherwise

## Credit

`one` started from the base of [ConvertOneNote2markdown](https://github.com/theohbrothers/ConvertOneNote2markdown), by

* [SjoerdV](https://github.com/SjoerdV)
* [nixsee](https://github.com/nixsee/)
* [theohbrothers](https://github.com/theohbrothers)

---

[Back to top](#onenote-exporter)
