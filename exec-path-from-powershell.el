;;; exec-path-from-powershell.el --- PowerShell support for exec-path-from-shell  -*- lexical-binding: t -*-

;; Copyright (C) 2026 Steve Purcell, Tim Felgentreff

;; Author: Tim Felgentreff <timfelgentreff@gmail.com>
;; Keywords: windows, environment
;; URL: https://github.com/purcell/exec-path-from-shell
;; Package-Requires: ((emacs "26.1"))

;; This file is not part of GNU Emacs.

;; This file is free software: you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This file is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this file.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; This file implements PowerShell support for `exec-path-from-shell'.

;;; Code:

(require 'cl-lib)
(require 'json)
(require 'subr-x)

(defvar exec-path-from-shell-debug)
(defvar exec-path-from-shell-variables)
(defvar exec-path-from-shell-warn-duration-millis)

(defgroup exec-path-from-powershell nil
  "PowerShell support for `exec-path-from-shell'."
  :prefix "exec-path-from-shell-powershell-"
  :group 'exec-path-from-shell)

(defcustom exec-path-from-shell-powershell-load-profile t
  "If non-nil, load the PowerShell profile."
  :type 'boolean
  :group 'exec-path-from-powershell)

(defcustom exec-path-from-shell-powershell-includes-visual-studio-environment t
  "If non-nil, also import the Visual Studio developer environment."
  :type 'boolean
  :group 'exec-path-from-powershell)

(defun exec-path-from-powershell--default-visual-studio-arch ()
  "Return the default Visual Studio architecture for the current machine."
  (let ((arch (downcase
               (or (getenv "PROCESSOR_ARCHITEW6432")
                   (getenv "PROCESSOR_ARCHITECTURE")
                   (car (split-string system-configuration "-"))
                   ""))))
    (cond
     ((member arch '("amd64" "x86_64" "x64")) "x64")
     ((member arch '("x86" "i386" "i486" "i586" "i686")) "x86")
     ((member arch '("arm64" "aarch64")) "arm64")
     ((string-equal arch "arm") "arm")
     (t "x64"))))

(defcustom exec-path-from-shell-powershell-visual-studio-arch
  (exec-path-from-powershell--default-visual-studio-arch)
  "Architecture argument for VsDevCmd or vcvarsall."
  :type 'string
  :group 'exec-path-from-powershell)

(defcustom exec-path-from-shell-powershell-visual-studio-host-arch
  (exec-path-from-powershell--default-visual-studio-arch)
  "Host architecture argument for VsDevCmd."
  :type 'string
  :group 'exec-path-from-powershell)

(defun exec-path-from-powershell--debug (msg &rest args)
  "Print MSG and ARGS like `message', but only if debug output is enabled."
  (when exec-path-from-shell-debug
    (apply #'message msg args)))

(defmacro exec-path-from-powershell--warn-duration (&rest body)
  "Evaluate BODY and warn if execution duration exceeds a time limit."
  (declare (indent 0))
  (let ((start-time (gensym))
        (duration-millis (gensym)))
    `(let ((,start-time (current-time)))
       (prog1
           (progn ,@body)
         (let ((,duration-millis (* 1000.0 (float-time (time-subtract (current-time) ,start-time)))))
           (if (> ,duration-millis exec-path-from-shell-warn-duration-millis)
               (message "Warning: exec-path-from-powershell execution took %dms" ,duration-millis)
             (exec-path-from-powershell--debug "PowerShell execution took %dms" ,duration-millis)))))))

(defun exec-path-from-powershell--ps-quote (s)
  "Quote S as a PowerShell single-quoted string."
  (concat "'" (replace-regexp-in-string "'" "''" s t t) "'"))

(defun exec-path-from-powershell--call-powershell (shell ps-script)
  "Run SHELL with PS-SCRIPT and return stdout as a string."
  (let* ((coding-system-for-read 'utf-8)
         (coding-system-for-write 'utf-8)
         (temp-ps1 (let ((inhibit-message t))
                     (make-temp-file "exec-path-from-powershell-" nil ".ps1" ps-script)))
         (args (append
                '("-ExecutionPolicy" "Bypass" "-NonInteractive")
                (unless exec-path-from-shell-powershell-load-profile
                  '("-NoProfile"))
                (list "-File" temp-ps1))))
    (unwind-protect
        (with-temp-buffer
          (exec-path-from-powershell--debug "Invoking %s with args %S" shell args)
          (exec-path-from-powershell--debug "PowerShell script contents:\n%s" ps-script)
          (let ((exit-code (exec-path-from-powershell--warn-duration
                             (apply #'call-process shell nil t nil args))))
            (unless (zerop exit-code)
              (error "PowerShell failed (exit %s): %s" exit-code (buffer-string))))
          (buffer-string))
      (ignore-errors (delete-file temp-ps1)))))

(defvar exec-path-from-powershell--visual-studio-variable-cache
  (make-hash-table :test #'equal)
  "Cache of discovered Visual Studio environment variable names.")

(defun exec-path-from-powershell--build-vs-block (env-table-var)
  "Return PowerShell code that merges a Visual Studio environment.
ENV-TABLE-VAR names the dictionary variable to update."
  (let ((arch (exec-path-from-powershell--ps-quote
               exec-path-from-shell-powershell-visual-studio-arch))
        (host (exec-path-from-powershell--ps-quote
               exec-path-from-shell-powershell-visual-studio-host-arch)))
    (concat
     "$vsOutput = @();"
     "try {"
     "  $progFiles = ${env:ProgramFiles(x86)};"
     "  if (-not $progFiles) { $progFiles = ${env:ProgramFiles} };"
     "  $vswhere = Join-Path $progFiles 'Microsoft Visual Studio\\Installer\\vswhere.exe';"
     "  if (-not (Test-Path $vswhere)) { throw \"vswhere not found: $vswhere\" };"
     "  $install = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath;"
     "  if (-not $install) { throw 'No suitable Visual Studio or Build Tools installation found.' };"
     "  $vsdev = Join-Path $install 'Common7\\Tools\\VsDevCmd.bat';"
     "  $vcvars = Join-Path $install 'VC\\Auxiliary\\Build\\vcvarsall.bat';"
     "  $cmdLine = $null;"
     "  if (Test-Path $vsdev) {"
     "    $cmdLine = ('call \"' + $vsdev + '\" -no_logo -arch=' + " arch " + ' -host_arch=' + " host ");"
     "  } elseif (Test-Path $vcvars) {"
     "    $cmdLine = ('call \"' + $vcvars + '\" ' + " arch ");"
     "  } else { throw \"Neither VsDevCmd.bat nor vcvarsall.bat found under: $install\" };"
     "  $cmd = 'chcp 65001>nul & ' + $cmdLine + ' & set';"
     "  $vsOutput = & cmd.exe /d /s /c $cmd;"
     "} catch {"
     "  throw $_"
     "};"
     "foreach ($line in $vsOutput) {"
     "  if ($line -match '^(.*?)=(.*)$') {"
     "    " env-table-var "[$matches[1]] = $matches[2]"
     "  }"
     "};")))

(defun exec-path-from-powershell--visual-studio-variable-cache-key (shell)
  "Return the cache key for the current Visual Studio environment settings.
SHELL is the PowerShell executable path."
  (list shell
        exec-path-from-shell-powershell-load-profile
        exec-path-from-shell-powershell-visual-studio-arch
        exec-path-from-shell-powershell-visual-studio-host-arch))

(defun exec-path-from-powershell--extract-json (raw)
  "Return the JSON payload embedded in RAW."
  (let ((start (string-match "[[{]" raw)))
    (unless start
      (error "Expected JSON from PowerShell, got: %S" raw))
    (substring raw start)))

(defun exec-path-from-powershell--visual-studio-variable-names (shell)
  "Return Visual Studio variable names that differ from the base shell."
  (let* ((cache-key (exec-path-from-powershell--visual-studio-variable-cache-key shell))
         (cached (gethash cache-key exec-path-from-powershell--visual-studio-variable-cache)))
    (or cached
        (let* ((script
                (concat
                 "$ErrorActionPreference='Stop';"
                 "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
                 "$baseEnv = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase);"
                 "[System.Environment]::GetEnvironmentVariables().GetEnumerator() | ForEach-Object { $baseEnv[$_.Key] = [string]$_.Value };"
                 "$vsEnv = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase);"
                 "$baseEnv.GetEnumerator() | ForEach-Object { $vsEnv[$_.Key] = $_.Value };"
                 (exec-path-from-powershell--build-vs-block "$vsEnv")
                 "$allNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase);"
                 "$baseEnv.Keys | ForEach-Object { [void]$allNames.Add($_) };"
                 "$vsEnv.Keys | ForEach-Object { [void]$allNames.Add($_) };"
                 "$result = foreach ($name in $allNames) {"
                 "  $baseValue = $baseEnv[$name];"
                 "  $vsValue = $vsEnv[$name];"
                 "  if ($baseValue -ne $vsValue) { $name }"
                 "};"
                 "$result | Sort-Object | ConvertTo-Json -Compress;"))
               (raw (exec-path-from-powershell--call-powershell shell script))
               (json-array-type 'list)
               (names (json-read-from-string (exec-path-from-powershell--extract-json raw))))
          (puthash cache-key names exec-path-from-powershell--visual-studio-variable-cache)
          names))))

(defun exec-path-from-powershell--maybe-extend-names-with-visual-studio-env (shell names)
  "Return NAMES augmented with Visual Studio variables when appropriate.
SHELL is the PowerShell executable path."
  (if (and exec-path-from-shell-powershell-includes-visual-studio-environment
           (equal names exec-path-from-shell-variables))
      (cl-remove-duplicates
       (append names (exec-path-from-powershell--visual-studio-variable-names shell))
       :test #'string-equal)
    names))

(defun exec-path-from-powershell--build-env-script (names include-vs)
  "Build a PowerShell script that returns JSON for NAMES.
When INCLUDE-VS is non-nil, merge the Visual Studio developer environment."
  (let* ((names-array
          (concat "@(" (mapconcat #'exec-path-from-powershell--ps-quote names ", ") ")"))
         (vs-block (and include-vs (exec-path-from-powershell--build-vs-block "$envTable"))))
    (concat
     "$ErrorActionPreference='Stop';"
     "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
     "$names=" names-array ";"
     "$envTable = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase);"
     "[System.Environment]::GetEnvironmentVariables().GetEnumerator() | ForEach-Object { $envTable[$_.Key] = [string]$_.Value };"
     "$result = @{};"
     "foreach ($name in $names) {"
     "  $result[$name] = $envTable[$name]"
     "};"
     (if vs-block
         (concat
          vs-block
          "foreach ($name in $names) {"
          "  $newValue = $envTable[$name];"
          "  $oldValue = $result[$name];"
          "  if (($null -ne $oldValue) -and ($null -ne $newValue) -and ($oldValue -ne $newValue) -and (($oldValue -match '.;.') -or ($newValue -match '.;.'))) {"
          "    $result[$name] = ($newValue + ';' + $oldValue)"
          "  } else {"
          "    $result[$name] = $newValue"
          "  }"
          "};")
       "")
     "$result | ConvertTo-Json -Compress;")))

(defun exec-path-from-powershell--evaluate-expressions (shell expressions)
  "Return a list of PowerShell evaluations for EXPRESSIONS using SHELL."
  (if (null expressions)
      '()
    (let* ((array-literal
            (concat "@(" (mapconcat #'exec-path-from-powershell--ps-quote expressions ", ") ")"))
           (script (concat
                    "$ErrorActionPreference='Stop';"
                    "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
                    "$exprs=" array-literal ";"
                    "$results = foreach ($expr in $exprs) {"
                    "  $value = Invoke-Expression $expr;"
                    "  if ($value -is [Array]) {"
                    "    $value = ($value -join [System.IO.Path]::PathSeparator)"
                    "  };"
                    "  if ($null -eq $value) {"
                    "    $null"
                    "  } else {"
                    "    [string]$value"
                    "  }"
                    "};"
                    "$results | ConvertTo-Json -Compress;"))
           (raw (exec-path-from-powershell--call-powershell shell script)))
      (let ((json-null :null)
            (json-array-type 'list))
        (mapcar (lambda (value)
                  (unless (eq value :null)
                    value))
                (json-read-from-string (exec-path-from-powershell--extract-json raw)))))))

(defun exec-path-from-powershell--decode-escapes (template)
  "Decode printf-style escapes in TEMPLATE."
  (let ((idx 0)
        (len (length template))
        (result (get-buffer-create " *exec-path-from-powershell-tmp*")))
    (unwind-protect
        (with-current-buffer result
          (erase-buffer)
          (while (< idx len)
            (let ((ch (aref template idx)))
              (setq idx (1+ idx))
              (cond
               ((/= ch ?\\)
                (insert-char ch 1))
               ((>= idx len)
                (insert-char ?\\ 1))
               (t
                (let ((esc (aref template idx)))
                  (setq idx (1+ idx))
                  (pcase esc
                    (?\\ (insert-char ?\\ 1))
                    (?\" (insert-char ?\" 1))
                    (?\' (insert-char ?\' 1))
                    (?a (insert-char ?\a 1))
                    (?b (insert-char ?\b 1))
                    (?e (insert-char ?\e 1))
                    (?f (insert-char ?\f 1))
                    (?n (insert-char ?\n 1))
                    (?r (insert-char ?\r 1))
                    (?t (insert-char ?\t 1))
                    (?v (insert-char ?\v 1))
                    (?0
                     (let ((start idx)
                           (digits ""))
                       (while (and (< idx len)
                                   (<= ?0 (aref template idx))
                                   (<= (aref template idx) ?7)
                                   (< (- idx start) 3))
                         (setq digits (concat digits (string (aref template idx))))
                         (setq idx (1+ idx)))
                       (insert-char (string-to-number
                                     (if (string-empty-p digits) "0" digits)
                                     8)
                                    1)))
                    (?x
                     (let ((start idx)
                           (digits ""))
                       (while (and (< idx len)
                                   (or (and (<= ?0 (aref template idx))
                                            (<= (aref template idx) ?9))
                                       (and (<= ?A (aref template idx))
                                            (<= (aref template idx) ?F))
                                       (and (<= ?a (aref template idx))
                                            (<= (aref template idx) ?f)))
                                   (< (- idx start) 2))
                         (setq digits (concat digits (string (aref template idx))))
                         (setq idx (1+ idx)))
                       (insert-char (string-to-number
                                     (if (string-empty-p digits) "0" digits)
                                     16)
                                    1)))
                    (_
                     (insert-char esc 1))))))))
          (buffer-string))
      (kill-buffer result))))

(defun exec-path-from-powershell-printf (shell str &optional args)
  "Return the PowerShell evaluation of STR formatted with ARGS using SHELL.
STR follows the same conventions as `exec-path-from-shell-printf'."
  (let* ((decoded (exec-path-from-powershell--decode-escapes str))
         (values (exec-path-from-powershell--evaluate-expressions shell args))
         (idx 0)
         (len (length decoded))
         (pieces (get-buffer-create " *exec-path-from-powershell-printf*")))
    (unwind-protect
        (with-current-buffer pieces
          (erase-buffer)
          (while (< idx len)
            (let ((ch (aref decoded idx)))
              (setq idx (1+ idx))
              (if (/= ch ?%)
                  (insert-char ch 1)
                (when (>= idx len)
                  (error "Incomplete format specifier in %S" str))
                (let ((next (aref decoded idx)))
                  (setq idx (1+ idx))
                  (pcase next
                    (?% (insert-char ?% 1))
                    (?s
                     (if (null values)
                         (error "Not enough arguments for format string")
                       (insert (or (pop values) ""))))
                    (_
                     (error "Unsupported format specifier %%%c" next)))))))
          (when values
            (dolist (extra values)
              (insert (or extra ""))))
          (buffer-string))
      (kill-buffer pieces))))

(defun exec-path-from-powershell-getenvs (shell names)
  "Get the environment variables with NAMES from PowerShell using SHELL.
The result is a list of (NAME . VALUE) pairs."
  (let* ((names (exec-path-from-powershell--maybe-extend-names-with-visual-studio-env
                 shell names))
         (script (exec-path-from-powershell--build-env-script
                  names
                  exec-path-from-shell-powershell-includes-visual-studio-environment))
         (raw (exec-path-from-powershell--call-powershell shell script))
         (json-object-type 'alist)
         (json-key-type 'string)
         (json-null nil))
    (json-read-from-string (exec-path-from-powershell--extract-json raw))))

(provide 'exec-path-from-powershell)

;;; exec-path-from-powershell.el ends here
