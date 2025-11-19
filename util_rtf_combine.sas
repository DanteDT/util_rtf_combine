/*** SOURCE:
     http://pharma-sas.com/a-sas-macro-to-combine-portrait-and-landscape-rtf-files-into-one-single-file/
   
   COMMENTS:
     Structure of SAS 9.4 RTF files at this time
     + OPENING SECTION
     + CONTENT SECTION - each page as a section, section breaks between
         first section:           \sectd\linex0\endnhere\pgwsxn15840\pghsxn12240\...
         remaining sections: \sect\sectd\linex0\endnhere\pgwsxn15840\pghsxn12240\...
         Pages in RTF = lines that start with "\sect"
     + CLOSING SECTION - "\pard}"
     + SAS seems to close NULL REPORTS with (see use of similar string in code, below):
         \pard}}{\*\bkmkstart IDX}{\*\bkmkend IDX}}
       These NULL REPORT contents need to be corrected, depending on where they
       appear in the COMBINED RTF, to ensure correct page display and numbering

   PARAMETERS:
   INPATH  - path to component SAS-generated RTF files
   OUTPATH - path to location for combined RTF file. SHOULD BE DIFFERENT from inpath
   OUTFILE - name of file for combined RTF pages.
   REGEX   - optional regex pattern for filtering filenames to include in combined RTF
             Note - Expecting an "m//" or "//" perl regex pattern
                   See: https://perldoc.perl.org/perlre.html#DESCRIPTION
                   See: https://go.documentation.sas.com/?cdcId=pgmsascdc&cdcVersion=9.4_3.4&docsetId=lefunctionsref&docsetTarget=p1vz3ljudbd756n19502acxazevk.htm&locale=en#n07tgr9r2iqygpn1c6qzv1rd0fhk
             NB: Regex pattern may need quoting by the calling program
             EG: /^t0[123]-[^9]/i
                 includes files like t01-... but not t01-9... 
                 (case INsensitive, like Win filenames)
   ONEPAGE - (OPTIONAL) Y/N to capture ONLY Page 1 of each component RTF. 
                        Default NO.
   VALIDATE- (OPTIONAL) Y/N to write out a validation ".txt" combination of original RTF contents.
                        Default YES.

   KNOWN ISSUE:
   DDT - SAS cannot read an RTF file that someone has open (macro alerts user)
         Work-around: Create a temp copy of the rtf, with suffix "_DEL.rtf"
                      Delete it once the combined file has been created.
   DDT - This macro only works to combine SAS-9.4-written RTF files with this structure:
   DDT - Each page is a Word "section" separated by section breaks (new page)
   DDT - Validate option (which preserves entire RTF source files) 
         cannot be used to validate &ONEPAGE reports

   MODIFICATIONS:
   ddt01 - fancy quote marks in web post to straight single or double quotes
   ddt02 - naming convention for TLFs, Windows path backslash
   ddt04 - hard-code page numbering since SAS 9.4 writes each page as a section,
           so we cannot combine RTFs and still use Page {PAGE} of {SECTIONPAGES},
           so we need to replace logical pages with hard-coded page numbers
     + RTF 1.5 specs: 
       https://www.biblioscape.com/rtf15_spec.htm
   ddt05 - remove a few uninitialized variables (never used)
   ddt06 - added validate=Y/N option to produce before/after text files for comparison
   ddt07 - attempt to detect and correct for NULL REPORTs as noted above
   ddt08 - ignore MS Word temp files, named "~$*.rtf"
   ddt09 - implement regex filename filter
***/

%macro util_rtf_combine(inpath=,
                        outpath=,
                        outfile=,
                        regex=,
                        onepage=n,
                        validate=y) 
       / minoperator;

%let max_infile_len = 32767;

%if "%substr(&regex,1,1)" NE "/" and
    "%substr(%upcase(&regex),1,1)" NE "M" %then %do;
  %goto quick_exit;
%end;

%let onepage = %substr(%upcase(&onepage),1,1);
%if "&onepage" = "Y" %then %let onepage = 1;
%else %let onepage = 0;

%let validate = %substr(%upcase(&validate),1,1);
%if "&validate" = "Y" %then %let validate = 1;
%else %let validate = 0;

*Get rtf file names from a folder;
*ddt - each INFILE with FILEVAR=, below, consumes 1 var. So make 2 up front;
  data rtffiles0 (keep=fileloc fileloc02 fnm);
    length fref $8 fnm fileloc fileloc02 $32767 ;

    rc = filename(fref, "&inpath");
    if rc = 0 then did = dopen(fref);

    dnum = dnum(did);
    putlog 'INFO: Number of files detected: ' dnum=;
    do i = 1 to dnum;
      fnm = dread(did, i);
      fid = mopen(did, fnm, 'i');

      if fid = 0 then 
         putlog 'WAR' 'NING: EXPECT TROUBLE. UNABLE to read file ' fnm;

      if fid > 0 and reverse(strip(upcase(fnm))) =: 'FTR.' then do;
        * ddt02 - Windows path backslash;
        fileloc=cats("&inpath\", fnm);
        fileloc02=fileloc;
        fnm = strip(tranwrd(fnm, '.rtf', ''));
        OUTPUT;
      end;
    end;
    rc = dclose(did);
  run;

  proc sql noprint;
    select max(length(fnm)), max(length(fileloc)), max(length(fileloc02)) 
           into
           :len_fnm trimmed, :len_fl trimmed, :len_fl2 trimmed
    from rtffiles0;
  quit;
  %put INFO: MAX filename and filepath LENGTHS are [&len_fnm &len_fl &len_fl2];

*Sort rtf files by TLF number;
  options varlenchk=nowarn;
  data rtffiles;
    length fnm       $%eval(2*&len_fnm) 
           fileloc   $%eval(2*&len_fl) 
           fileloc02 $%eval(2*&len_fl2) ;
    set rtffiles0;

    * ddt08 - Ignore temp MS Word file-lock files ;
    if fnm =: '~$' then delete;

    %* d09 - Optional regex filename filter ;
    %if %length(%superq(regex)) > 0 %then %do;
      if prxmatch(%sysfunc(quote(&regex)), fnm) = 0 then delete;
    %end;

    * ddt02 - naming conventions for TLFs;
    if lowcase(fnm) =: 't' then ord = 1;
    else if lowcase(fnm) =: 'l' then ord = 2;
    else if lowcase(fnm) =: 'f' then ord = 3;
    else ord = 4;
  run;

  proc sort data = rtffiles; 
    by ord fnm; 
  run;
  options varlenchk=warn;

*DDT06 - VALIDATE - create validation text output, if requested, in same order;
  %if &validate %then %do;
    options noxwait xsync;

    data _null_;
      set rtffiles;

      if _n_ = 1 then redir = ' >';
      else redir = '>>';

      length cmd $1000;
      cmd = "type "||catx(' ', fileloc, redir)||" &outpath.\validate_&outfile..rtf.txt";
      call system(cmd);
    run;
  %end;

*Create macro variable which contains all rtf files;
  proc sql noprint;
    select quote(strip(fileloc)) into :rtffiles separated by ', '
    from rtffiles;
  quit;

*Create filename rtffiles ###;
* DDT TODO - confusing to overload RTFFILES name ###;
*          - may not be necessary. {fref} is a placeholder, below, when INFILE {fref}... FILEVAR= specified ###;
*          - could use fileref dummy "dummy" ###;
  filename rtffiles (&rtffiles); 

*ddt04 - Count pages in each file - for Page X of Y replacement;
  data rtffiles;
    set rtffiles end=no_more;
    length rtfcode $&max_infile_len;
    drop rtfcode;
    
    retain numpages 0;
    numpages = 0;

    do until (eof);
      *Read next RTF files when fileloc changes. Consume 1 fileloc var*;
      * - see documentation for INFILE FILEVAR= option - *;
      infile rtffiles lrecl=&max_infile_len end=eof filevar=fileloc02;
      INPUT;
      rtfcode = strip(_infile_);

      if rtfcode =: '\sect' then numpages+1;
    end;
  run;


*Data START - Collect document opening content from FIRST rtf file;
  data _null_;
    set rtffiles;
    call symput('rtf_file_01', strip(fileloc));
    STOP;
  run;

  data start;
    length rtfcode $&max_infile_len;
    infile "&rtf_file_01" lrecl = &max_infile_len end = eof;
    INPUT ;

    *ddt - Warn if infile line is potentially truncated, then remove leading/trailing blanks;
    rtfcode = _infile_;
    if length(_infile_) = &max_infile_len then
       putlog 'WAR' "NING (util_rtf_combine): MAX LENGTH OF _INFILE_ line (&max_infile_len) - potential truncation.";
    rtfcode = strip(rtfcode);

    *ddt - Opening section stops with first content section;
    if rtfcode =: '\sectd' then 
       STOP;
  run;

/***Data RTF - Collect SECTION CONTENT for each rtf file
    SAS 9.4 writes each page as a SECTION, using Page X or Y for overall doc!
    To preserve Page X of Y within each RTF table, 
    we must replace page fields with hard-coded pages
***/
  data rtf (keep = rtfcode);
    length rtfcode $&max_infile_len ;
    set rtffiles end=last;

    retain total_combined_rtfs total_combined_pages firstsect sof kpfl 0;
    sof = 1;  * Start of next rtf file *;
    kpfl = 0; * Flag to keep content, once CONTENT SECTION starts *;

    total_combined_rtfs + 1;
    total_combined_pages + numpages;
    if last then do;
      call symput('total_combined_rtfs', strip(put(total_combined_rtfs, best10.-L)));
      call symput('total_combined_pages', strip(put(total_combined_pages, best10.-L)));
    end;

    do until (eof);
      *change of filevar value (fileloc from rtf) forces opening next rtffile;
      infile rtffiles lrecl=&max_infile_len end=eof filevar=fileloc;
      INPUT;
      rtfcode=_infile_;
      rtfcode=strip(rtfcode);

      /*** Skip OPENING SECTION from each RTF, just keep SECTION CONTENT, but as one single SECTION
           -- PAGE BREAKs between pages, rather than section breaks
           Remove RTF header section and replace first "\sectd" with "\sect\sectd\..."
           FOCUS here on image of rtf structure from: 
             http://pharma-sas.com/a-sas-macro-to-combine-portrait-and-landscape-rtf-files-into-one-single-file/
      ***/
      if sof then kpfl=0;

      *--- FIRST SECTION detected, after opening header ;
      if rtfcode =: '\sect' then do;
        kpfl+1;
        if rtfcode =: '\sectd' then do;
          if not firstsect then firstsect+1;
          else rtfcode = '\sect' || strip(rtfcode);
        end;
      end;

  * Deassign placeholder fileref ###;
  filename rtffiles; 

/*      *--- EXIT before writing page 2, if combining only page 1s ;*/
/*      if kpfl > 1 and &onepage then leave;*/

      *Hard-code Page X of Y;
      if index(rtfcode, '{\field{\*\fldinst { PAGE }}}') or 
         index(rtfcode, '{\field{\*\fldinst { NUMPAGES }}}') then do;

        length this_pagenum this_rtf_numpages $10;
        this_pagenum      = put(kpfl,best10.-L);
        this_rtf_numpages = put(numpages,best10.-L);
        putlog 'NOTE: (util_rtf_combine) Fixing page numbering to: Page ' this_pagenum 'of ' this_rtf_numpages;

        if index(rtfcode, '{\field{\*\fldinst { PAGE }}}') then 
           rtfcode = tranwrd(rtfcode, '{\field{\*\fldinst { PAGE }}}', strip(this_pagenum));
        if index(rtfcode, '{\field{\*\fldinst { NUMPAGES }}}') then 
           rtfcode = tranwrd(rtfcode, '{\field{\*\fldinst { NUMPAGES }}}', strip(this_rtf_numpages));

      end;

      *Remove RTF closing } except for last file;
      if eof and not last then do;

        *--- ddt07 - ATTEMPT TO CORRECT FOR NULL REPORT, which corrupts paging of COMBINED RTF ;
        if rtfcode = '\pard}}{\*\bkmkstart IDX}{\*\bkmkend IDX}}' then do;
          put 'WAR' 'NING: (util_rtf_combine) NULL REPORT detected - Adding note to blank page for ' fnm= rtfcode=;
          *--- NOTE - DO NOT CLOSE THE RTF document with final brace, as done further down! ;
          rtfcode = '\pard}}{\*\bkmkstart IDX}{\*\bkmkend IDX}\par NULL REPORT detected' ;
        end;
        else do;
          put 'NOTE: (util_rtf_combine) Converting file closer to "\pard" from: ' fnm= rtfcode=;
          rtfcode = '\pard';
        end;
      end;

      *Output concatenated rtf ;
      if kpfl then do;
        *--- ddt07 - ATTEMPT TO CORRECT FOR NULL REPORT, which corrupts paging of COMBINED RTF ;
        if rtfcode = '\pard}}{\*\bkmkstart IDX}{\*\bkmkend IDX}}' then do;
          put 'WAR' 'NING: (util_rtf_combine) NULL REPORT detected - Adding note to blank page for ' fnm= rtfcode=;
          rtfcode = '\pard}}{\*\bkmkstart IDX}{\*\bkmkend IDX}\par NULL REPORT detected}' ;
        end;

        *--- ONEPAGE? Only write out page 1s ;
        if not &onepage or (&onepage and kpfl = 1) then
           OUTPUT;
      end;

      sof=0;
    end;

    *--- Having written page one ONLY, close this section (or the entire doc) ;
    if &onepage and last then do;
      rtfcode = '\pard}';
      OUTPUT;
    end;
  run;

  data _null_;
    * ddt02 - Windows path backslash;
    file "&outpath\&outfile..rtf" lrecl=&max_infile_len nopad;

    *concatenate rtf header, document info, all rtf content sections;
    set start rtf;
    put rtfcode;
  run;

  %put NOTE: (util_rtf_combine) Combined RTF file is: &outpath.\&outfile..rtf;
  %put NOTE: (util_rtf_combine) Total combined RTFs: &total_combined_rtfs;
  %if &onepage %then 
      %put NOTE: (util_rtf_combine) Total pages in combined RTF file should be: &total_combined_rtfs;
  %else
      %put NOTE: (util_rtf_combine) Total pages in combined RTF file should be: &total_combined_pages;

%QUICK_EXIT:

%mend util_rtf_combine;
