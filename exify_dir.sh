#!/bin/sh

#  -----------------------------------------------
#  Parameters controlling input and output sources
#  -----------------------------------------------

sample_dir="${PWD}/samples" # directory with the sample Office Open XML files
output_file="${PWD}/data/ooxml_results.csv"   # where to write the test results?

# ------------------
# EXI codec location
# ------------------

EXIFICIENT="java -jar ${PWD}/exip/exip.jar"

#  --------------------------------------------------
#  Erase any existing output file and start a new one
#  --------------------------------------------------

rm ${output_file}
touch ${output_file}
echo "file,variable,value" >> ${output_file}

#  ---------------------
#  Iterate through files
#  ---------------------

for f in `find ${sample_dir} -type f \( -name "*.docx" -or -name "*.xlsx" -or -name "*.pptx" \)`
do
    

    path=${f%/*.*}/    # Break apart the current file location into path,filename, and extension
    base=$(basename ${f})
    extension=${base##*.}
    filename=${base%.*}


    echo "Processing ${f}..."            
            
    # -------------------------------
    # Write uncompressed size to file
    # -------------------------------

    eval $(stat -s ${path}/${filename}.${extension})
    echo "${filename},${extension},$st_size" >> ${output_file}

    # -----------------------------------------------------
    # Capture the file extension and convert for rezipping. 
    # It is converted like so:
    #   docx --> doce, using wml.xsd
    #   xlsx --> xlse, using sml.xsd
    #   pptx --> ppte, using pml.xsd
    # -----------------------------------------------------

    case ${extension} in
        docx)
            schema="${PWD}/xsd/ooxml-strict/wml.xsd"
            extension='doce'
            echo "Converting docx to ${extension}"
            ;;
        xlsx)
            schema="${PWD}/xsd/ooxml-strict/sml.xsd"
            extension='xlse'
            echo "Converting xlsx to ${extension}"
            ;;
        pptx)
            schema="${PWD}/xsd/ooxml-strict/pml.xsd"
            extension='ppte'
            echo "Converting pptx to ${extension}"
            ;;
    esac

    # --------------------------------------
    # Unzip the file into a temporary folder
    # --------------------------------------
    
    mkdir temp
    cp ${f} temp
    cd temp
    unzip $(basename ${f})
    rm $(basename ${f})
     
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

    zip -r ${path}/${filename}.${extension} *
    cd ..

    # -----------------------------
    # Write compressed size to file
    # -----------------------------

    eval $(stat -s ${path}/${filename}.${extension})
    echo "${filename},${extension},$st_size" >> ${output_file}
    
    # ---------------------
    # Trash the temp folder
    # ---------------------

    rm -r temp

done
