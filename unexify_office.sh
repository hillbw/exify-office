#!/bin/sh

# -------------------------------------------------------------
#  * Script to restore an exified Microsoft Office XML file
#  * Accepts 1 argument: a path to an .doce, .ppte, or .xlse file
#  * Resulting decompressed version is placed in same directory 
#    with '_unexify' appended to the filename.
# -------------------------------------------------------------


# ------------------
# EXI codec location
# ------------------

EXIFICIENT="java -jar ${PWD}/exip/exip.jar"

# --------------------------------------
# Verify file extension & file existence
# --------------------------------------

target=$1
ext_regex='^(.*)\.(doce|xlse|ppte)$'

if [[ ${target} =~ ${ext_regex} ]]; then 
	if [ -f ${target} ]; then

		# -----------------------------------------------------
		# Capture the file extension and convert for rezipping. 
		# It is converted like so:
		#   doce --> docx, using wml.xsd
		#   xlse --> xlsx, using sml.xsd
		#   ppte --> pptx, using pml.xsd
		# -----------------------------------------------------
		
		path=${BASH_REMATCH[1]}
		ext=${BASH_REMATCH[2]}

		case ${ext} in
        	doce)
							schema="${PWD}/xsd/ooxml-strict/wml.xsd"
        	    ext='docx'
        	    echo "Converting doce to ${ext}"
        	    ;;
	        xlse)
							schema="${PWD}/xsd/ooxml-strict/sml.xsd"
	            ext='xlsx'
	            echo "Converting xlse to ${ext}"
	            ;;
	        ppte)
							schema="${PWD}/xsd/ooxml-strict/pml.xsd"
	            ext='pptx'
	            echo "Converting ppte to ${ext}"
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

		for f in `find . -type f -name "*.exi"`
		do	

			# -------------------------
			# Strip .exi file extension
			# -------------------------

			strip_ext_regex='^(.*)\.exi$'
			[[ ${f} =~ ${strip_ext_regex} ]]
			f_new=${BASH_REMATCH[1]}

			# ---------------------------------------
			# Decompress with EXI in compression mode
			# ---------------------------------------

			${EXIFICIENT} -exi_in ${f} -xml_out ${f_new} -compression -preserve_prefixes -schema ${schema}
			
			# -----------------------------
			# Trash the compressed EXI file
			# -----------------------------
			rm ${f}

		done

		# --------------------------------------
		# Rezip the file with original extension
		# --------------------------------------

		zip -r ../${path}_unexify.${ext} *
		cd ..

		# ---------------------
		# Trash the temp folder
		# ---------------------

		rm -r temp

	fi
fi
