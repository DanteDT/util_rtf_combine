# util_rtf_combine

Thanks to @rogerjdeangelis for [preserving the original macro](https://github.com/rogerjdeangelis/utl-sas-macro-to-combine-rtf-files-into-one-single-file) from the now-defunct pharma-sas.com site. That is the source of this fork.

In case it helps others, I extended that original:
- file-mask for fine-tuning RTF file selection
- page-numbering in combined RTF
- validation option, to support text-based compare of changes made to combined content vs. original RTFs
- handle individual null reports
- optionally create a "page 1s" combined RTF - only page 1 of each contributing RTF
