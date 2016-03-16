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
# Parse macro content, extract function usage and add logging code to them

import sys
import re

class vhook:
	IMPORTANT_FUNCTIONS_LIST = ["CallByName"]
	EXTERNAL_FUNCTION_DECLARATION_REGEXP = re.compile(r'Declare *(?:PtrSafe)? *(?:Sub|Function) *(.*?) *Lib *"[^"]+" *(?:Alias *"([^"]+)")?')
	EXTERNAL_FUNCTION_REGEXP = None
	EXTERNAL_FUNCTION_REGEXP_2 = None
	BEGIN_FUNCTION_REGEXP = re.compile(r"\s*(?:function|sub) (.*?)\(", re.IGNORECASE)
	END_FUNCTION_REGEXP = re.compile(r"^\s*end\s*(?:function|sub)", re.IGNORECASE)
	METHOD_CALL_REGEXP = re.compile(r"^\s*([a-z_\.0-9]+\.[a-z_0-9]+)\s*\((.*)\)", re.IGNORECASE)
	METHOD_CALL_REGEXP_2 = re.compile(r"^\s*([a-z_\.0-9]+\.[a-z_0-9]+) +(.*)", re.IGNORECASE)
	IMPORTANT_FUNCTION_REGEXP = re.compile(r"^(Set\s+)?([a-z0-9]+\s*=)?\s*("+"|".join(IMPORTANT_FUNCTIONS_LIST)+")\s*\((.*)\)", re.IGNORECASE)
	IMPORTANT_FUNCTION_REGEXP_2 = re.compile(r"^(Set\s+)?([a-z0-9]+\s*=)?\s*("+"|".join(IMPORTANT_FUNCTIONS_LIST)+")(.*)", re.IGNORECASE)

	is_auto_open_function = False
	current_function_name = ""
	declared_function_original_names = {}

	lines = [];
	i = 0
	counter = 0
	line = ""

	output = []


	def print_info(self, message):
		print message

	def __init__(self):
		self.lines = sys.stdin.read()

		if len(self.lines) == 0:
			self.print_info("[-] Missing input")
			sys.exit()

		# For external functions we can ignore long lines errors
		self.prepare_external_function_calls(self.lines.replace(" _\n", " "))
		self.lines = self.lines.split("\n")
		self.dispatch()

	def add_content_to_output(self, content):
		self.output.append(content)

	def add_line_to_output(self, i):
		self.add_content_to_output(self.get_line(i))

	def add_current_line_to_output(self):
		self.add_content_to_output(self.get_current_line())


	"""
		See: https://msdn.microsoft.com/en-us/library/ba9sxbw4.aspx
		We simply skip those kind of lines
	"""
	def is_long_line(self):
		counter = 0
		while True:
			if self.get_line(self.i+counter)[-2:] == " _":
				counter += 1
			else:
				return counter

	"""
		Those functions starts automatically when a document is opened
		We need to initialize our logger there
	"""
	def is_autostart_function(self):
		auto_open_list = ["document_open", "workbook_open", "autoopen", "auto_open"]

		if any(x in self.current_function_name.lower() for x in auto_open_list):
			return True

		return False

	"""
		If current line is function declaration, get its name
		Its then used for checking if we return something from this function
	"""
	def is_begin_function_line(self):
		matched = re.search(self.BEGIN_FUNCTION_REGEXP, self.get_current_line())
		if matched and not "declare" in self.get_current_line().lower():
			self.current_function_name = matched.group(1).strip()
			return True

		return False

	"""
		Check if its function end
		We add there exception handler
	"""
	def is_end_function_line(self):
		if re.search(self.END_FUNCTION_REGEXP, self.get_current_line()):
			return True

		return False

	"""
		For non-object return types, you assign the value to the name of function
		So we can check if this function return something and add our logger here
	"""
	def is_return_string_from_function_line(self):
		if self.current_function_name != "":
			if re.search("^ *"+re.escape(self.current_function_name)+" *=", self.get_current_line(), re.IGNORECASE):
				return True
		return False

	"""
		Found all `Class.Method params` and Class.Method (params)` calls
		We dont support params passed by name, like name:=value
		We need to check if its not method assign like variable = Class.Method
	"""
	def is_method_call_line(self):
		matched = re.search(self.METHOD_CALL_REGEXP, self.get_current_line())
		if not matched:
			matched = re.search(self.METHOD_CALL_REGEXP_2, self.get_current_line())

		if matched:
			method_name = matched.group(1).strip()
			params = matched.group(2).strip()

			if len(params) > 0:
				# We dont support params passed by name
				# And Class.Method = 1
				if "=" in self.get_current_line():
				 	return False

				return [method_name, params]

			return [method_name]
		return False

	"""
		Check if its call to previously defined external library
	"""
	def is_external_function_call_line(self):
		# Do we have any external declarations
		if self.EXTERNAL_FUNCTION_REGEXP == None:
			return False

		# Skip if its declaration, not usage
		if re.search(self.EXTERNAL_FUNCTION_DECLARATION_REGEXP, self.get_current_line()):
			return False

		matched = re.search(self.EXTERNAL_FUNCTION_REGEXP, self.get_current_line())

		if not matched:
			matched = re.search(self.EXTERNAL_FUNCTION_REGEXP_2, self.get_current_line())

		if matched:
			name = matched.group(1).strip()
			rest = matched.group(2).strip()
			if name in self.declared_function_original_names:
				rest = re.sub(r"ByVal", "", rest, flags=re.IGNORECASE)
				return [name, rest]

		return False

	"""
		We hook some important function like CallByName which cannot be hooked using another techniques
	"""
	def is_important_function_call_line(self):
		matched = re.search(self.IMPORTANT_FUNCTION_REGEXP, self.get_current_line())

		if not matched:
			matched = re.search(self.IMPORTANT_FUNCTION_REGEXP_2, self.get_current_line())

		if matched:
			name = matched.group(3).strip()
			rest = matched.group(4).strip()

			return [name, rest]

		return False

	"""
		Find all external library declarations like:
		Private Declare Function GetDesktopWindow Lib "user32" () As Long
	"""
	def prepare_external_function_calls(self, content):
		declared_function_list = []

		for f in re.findall(self.EXTERNAL_FUNCTION_DECLARATION_REGEXP, content):
			declared_function_list.append(re.escape(f[0].strip()))
			if f[1] != "":
				self.declared_function_original_names[f[0].strip()] = f[1].strip()
			else:
				self.declared_function_original_names[f[0].strip()] = f[0].strip()

		if len(declared_function_list) > 0:
			self.print_info("[+] Found external function declarations: {}".format(",".join(self.declared_function_original_names.values())))

			self.EXTERNAL_FUNCTION_REGEXP  = re.compile("({})\s*\((.*)\)".format("|".join(declared_function_list)))
			self.EXTERNAL_FUNCTION_REGEXP_2  = re.compile("({})\s*(.*)".format("|".join(declared_function_list)))

	"""
		Get single line by ids number
	"""
	def get_line(self, i):
		if i < self.counter:
			return self.lines[i]
		return ""

	"""
		Get current line
	"""
	def get_current_line(self):
		if self.i < self.counter:
			return self.lines[self.i]
		return ""

	"""
		Set current line, so we can then use get_current_line
	"""
	def set_current_line(self):
		self.line = self.get_line(self.i)

	"""
		Some function have special aliases for null support
	"""
	def replace_function_aliases(self):
		line = self.lines[self.i]
		line = re.sub(r"(VBA\.CreateObject)", "CreateObject", line, flags=re.IGNORECASE)
		line = re.sub(r"Left\$", "Left", line, flags=re.IGNORECASE)
		line = re.sub(r"Right\$", "Right", line, flags=re.IGNORECASE)
		line = re.sub(r"Mid\$", "Mid", line, flags=re.IGNORECASE)
		line = re.sub(r"Environ\$", "Environ", line, flags=re.IGNORECASE)
		self.lines[self.i] = line

	"""
		Main program loop
	"""
	def dispatch(self):
		self.i = 0
		self.counter = len(self.lines)
		while self.i < self.counter:
			self.set_current_line()
			self.replace_function_aliases()

			is_long_line = self.is_long_line()
			is_method_call_line = self.is_method_call_line()
			is_external_function_call_line = self.is_external_function_call_line()
			is_important_function_call_line = self.is_important_function_call_line()

			if is_long_line > 0:
				self.print_info("[+] Found long line, skip {} lines".format(is_long_line))

				for ii in range(self.i, self.i+is_long_line+1):
					self.add_line_to_output(ii)

				self.i += is_long_line+1
				continue
			elif self.is_begin_function_line():
				if self.is_autostart_function():
					self.print_info("[+] Found autostart function - {}".format(self.current_function_name))

					self.is_auto_open_function = True

					self.add_current_line_to_output()
					self.add_content_to_output("On Error GoTo vhook_exception_handler:")
					self.add_content_to_output("vhook_init")
				else:
					self.print_info("[+] Found function - {}".format(self.current_function_name))

					self.is_auto_open_function = False

					self.add_current_line_to_output()
			elif self.is_return_string_from_function_line():
				self.print_info("\t[+] Function return string")

				self.add_current_line_to_output()
				self.add_content_to_output('log_return_from_string_function "{}", {}'.format(self.current_function_name, self.current_function_name))
			elif is_method_call_line != False:
				if len(is_method_call_line) == 1:
					self.print_info("\t[+] Found call to method: {}".format(is_method_call_line[0]))
					self.add_content_to_output('log_call_to_method "{}"'.format(is_method_call_line[0]))
				else:
					self.print_info("\t[+] Found call to method with params: {} - {}".format(is_method_call_line[0], is_method_call_line[1]))
					self.add_content_to_output("log_call_to_method \"{}\", {}".format(is_method_call_line[0], is_method_call_line[1]))
				self.add_current_line_to_output()
			elif is_external_function_call_line != False:
				self.print_info("\t[+] Found externall call to {}".format(self.declared_function_original_names[is_external_function_call_line[0]]))

				self.add_content_to_output('log_call_to_function "{}", {}'.format(self.declared_function_original_names[is_external_function_call_line[0]], is_external_function_call_line[1]))
				self.add_current_line_to_output()
			elif is_important_function_call_line != False:
				self.print_info("\t[+] Found important function {}".format(is_important_function_call_line[0]))

				self.add_content_to_output('log_call_to_function "{}", {}'.format(is_important_function_call_line[0], is_important_function_call_line[1]))
				self.add_current_line_to_output()
			elif self.is_end_function_line():
				if self.is_auto_open_function:
					self.print_info("\t[+] Add exception handler")
					self.add_content_to_output('vhook_exception_handler:')
					self.add_content_to_output('vhook_log ("Exception: " & Err.Description)')
					self.add_content_to_output('On Error Resume Next')

				self.add_current_line_to_output()
			else:
				self.add_current_line_to_output()

			self.i += 1
		print "|*&|VHOOK_SPLITTER|&*|"
		print "\n".join(self.output)

vhook()
