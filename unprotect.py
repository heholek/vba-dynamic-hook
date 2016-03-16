# Code related to ESET's VBA Dynamic Hook research
# For feedback or questions contact us at: github@eset.com
# https://github.com/eset/vba-dynamic-hook/
#
# This code is provided to the community under the two-clause BSD license as
# follows:
#
# Copyright (C) 2016 ESET
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice, this
# list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
# this list of conditions and the following disclaimer in the documentation
# and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Kacper Szurek <kacper.szurek@eset.com>
#
# Try to remove VBA password protection from `.doc` file

import sys, re, zipfile, os, shutil

if len(sys.argv) != 3:
	print "[-] Usage: unprotect.py [file_with_password] [new_file_without_password]"
	sys.exit(1)

abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

def update_zip(old_file_name, new_file_name, filename, data):
	with zipfile.ZipFile(old_file_name, 'r') as zin:
		with zipfile.ZipFile(new_file_name, 'w') as zout:
			for item in zin.infolist():
				if item.filename != filename:
					zout.writestr(item, zin.read(item.filename))
			zout.writestr(filename, data)

REPLACE_STRING = 'Description='

old_file_name = sys.argv[1]
new_file_name = sys.argv[2]
is_zip = False

if open(old_file_name, "rb").read(2) == "PK":
	if not zipfile.is_zipfile(old_file_name):
		print "[-] Not zip file, probably corrupted docm file"
		sys.exit(1)

	is_zip = True
	with zipfile.ZipFile(old_file_name, 'a') as archive:
		name_list = archive.namelist()

		if "word/vbaProject.bin" not in name_list:
			print "[-] Cannot find vbaProject"
			sys.exit(1)

		content = archive.read("word/vbaProject.bin")
else:
	with open(old_file_name, "rb") as f:
		content = f.read()

match = re.search(r'(CMG="[A-Za-z0-9]+"\s*DPB="[A-Za-z0-9]+"\s*GC="[A-Za-z0-9]+")', content)
if not match:
	print "[+] Try advanced method"
	match = re.search(r'(GC="[A-Za-z0-9]+")', content)

if match:
	new_content = ""
	pass_len = len(match.group())
	new_content = content[:match.start()]
	new_content += REPLACE_STRING
	new_content += " " * (pass_len - len(REPLACE_STRING))
	new_content += content[match.start()+pass_len:]

	if is_zip:
		update_zip(old_file_name, new_file_name, "word/vbaProject.bin", new_content)
	else:
		with open(new_file_name, "wb") as f:
			f.write(new_content)

	print "[+] Password removed"
else:
	shutil.copyfile(old_file_name, new_file_name)
	print "[-] Probably without pasword"
