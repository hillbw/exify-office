#!/bin/sh

# ----------------------------------------------------------------
#  * Script to unzip, exify, and rezip a Microsoft Office XML file
#  * Accepts 1 argument: a path to .docx, .xlsx, or .pptx file
#  * Resulting compressed version is placed in same directory
# ----------------------------------------------------------------


# ------------------
# EXI codec location
# ------------------

EXIFICIENT="java -jar ${PWD}/exip/exip.jar"

# --------------------------------------
# Verify file extension & file existence
# Use regex to capture extension
# --------------------------------------

target=$1
ext_regex='^(.*)\.(docx|xlsx|pptx)$'

if [[ ${target} =~ ${ext_regex} ]]; then 
	if [ -f ${target} ]; then

		# -----------------------------------------------------
		# Capture the file extension and convert for rezipping. 
		# It is converted like so:
		#   docx --> doce, using wml.xsd
		#   xlsx --> xlse, using sml.xsd
		#   pptx --> ppte, using pml.xsd
		# -----------------------------------------------------
		
		path=${BASH_REMATCH[1]}
		ext=${BASH_REMATCH[2]}

		echo path:  ${path}
		echo ext:   ${ext}

		case ${ext} in
        	docx)
							schema="${PWD}/xsd/ooxml-strict/wml.xsd"
        	    ext='doce'
        	    echo "Converting docx to ${ext}"
        	    ;;
	        xlsx)
							schema="${PWD}/xsd/ooxml-strict/sml.xsd"
	            ext='xlse'
	            echo "Converting xlsx to ${ext}"
	            ;;
	        pptx)
							schema="${PWD}/xsd/ooxml-strict/pml.xsd"
	            ext='ppte'
	            echo "Converting pptx to ${ext}"
	            ;;
	    esac

		# --------------------------------------
		# Unzip the file into a temporary folder
		# --------------------------------------
		
		mkdir temp
		cp ${target} temp
		cd temp
		unzip $(basename ${target})
		rm $(basename ${target})
		 
		# -------------------------------------------------
		# Iterate through all XML files in temporary folder
		# -------------------------------------------------

		for f in `find . -type f -name "*.rels" -or -name "*.xml"`
		do	

			# ----------------------------------
			# Compress with EXI in compress mode
			# ----------------------------------

			${EXIFICIENT} -xml_in ${f} -exi_out ${f}.exi -compression -preserve_prefixes -schema ${schema}
			
			# --------------------------
			# Trash the uncompressed XML
			# --------------------------
		
			rm ${f}

		done

		# ---------------------------------
		# Rezip the file with new extension
		# ---------------------------------
		
		zip -r ../${path}.${ext} *
		cd ..

		# --------------------
		# Dump a status report
		# --------------------
		echo "\nStatus Report"
		echo "\n-------------\n"
		ls -l ${path}.*
		
		
		# ---------------------
		# Trash the temp folder
		# ---------------------

		rm -r temp



	fi
fi
