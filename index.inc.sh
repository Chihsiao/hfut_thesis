# shellcheck shell=bash

import_module office_word
import_module docx_patcher
import_module pandoc_crossref
import_module abstract_section

PAPYRUS_YPP_FLAGS+=(
    -l "$__MODULE_ROOT__/index.lua"
)

PAPYRUS_PANDOC_FLAGS+=(
    --template="$__MODULE_ROOT__/template.xml"
)

__MODULE_ROOT_hfut_thesis__="$__MODULE_ROOT__"

function __hfut_thesis_setup() {
    local ref_time ref_file="${target_dir:?}/$PAPYRUS_BASENAME/ref.docx"
    local scm_file="$__MODULE_ROOT_hfut_thesis__/thesis.scm.sh"
    PAPYRUS_PANDOC_FLAGS+=(--reference-doc="$ref_file")

    if ref_time="$(date -r "$ref_file" +%s 2> /dev/null)"; then
        if [[ "$ref_time" -ge "$(date -r "$scm_file" +%s)" ]]; then
            return 0
        fi
    fi

    mkdir -p -- "$(dirname -- "$ref_file")"
    docx-patcher.sh "$scm_file" "$OOXML_PATCHER_ROOT/examples/reference.docx" "$ref_file"
}

PAPYRUS_INITIALIZERS+=(
    __hfut_thesis_setup
)

function __hfut_thesis_postprocess() {
    docx-patcher.sh "$__MODULE_ROOT_hfut_thesis__/patch.sh" \
        "${bundle_file:?}" "$bundle_file"
}

PAPYRUS_POSTPROCESSORS+=(
    __hfut_thesis_postprocess
)
