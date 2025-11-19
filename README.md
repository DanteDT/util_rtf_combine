# util_rtf_combine

Thanks to [@rogerjdeangelis](https://github.com/rogerjdeangelis) for [preserving the original macro](https://github.com/rogerjdeangelis/utl-sas-macro-to-combine-rtf-files-into-one-single-file) from the now-defunct pharma-sas.com site. That is the source of this fork.

In case it helps others, I extended that original :
- file-mask for fine-tuning RTF file selection
- page-numbering in combined RTF
- validation option, to support text-based compare of changes made to combined content vs. original RTFs
- handle individual null reports
- optionally create a "page 1s" combined RTF - only page 1 of each contributing RTF

See [util_rtf_combine.sas](util_rtf_combine.sas) modified version of pharma-sas.com original [utl_rtfcombine.sas](utl_rtfcombine.sas)

Knows issues:
- Windows-style backslashes in the code, since I was working in a Windows env at the time.

References:
- https://www.lexjansen.com/pharmasug/2010/PO/PO05.pdf
- https://www.lexjansen.com/pharmasug-cn/2015/DV/PharmaSUG-China-2015-DV28.pdf
- https://learn.microsoft.com/en-us/previous-versions/office/developer/office2000/aa140302(v=office.10)
