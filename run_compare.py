from transform_data.edit_dataset import edit_data_file
from distribution.divide_dataset import split_dataset
from genotype_evaluation.genotype_scorer import create_genotype_matrix
from formatted_matrices.matrix_formatter import format_trait_matrix
from reordered_matrices.column_shifter import reorder_marker_columns
from evaluated_status_matrices.gene_status_evaluator import evaluate_gene_status
from annotated_matrices.name_mapper import insert_genotype_names
from phenotype_evaluation.phenotype import phenotype_function






def main():
    genotypes_data = "IMPORT/Genotype.xlsx"
    trait_specific_data = "IMPORT/Trait-specific.xlsx"

    edit_data_file(genotypes_data,trait_specific_data)
    split_dataset("transform_data/Trait_specific_updated.xlsx")

    create_genotype_matrix("transform_data/Genotypes_targeted_updated.xlsx", "distribution/Same_Position_Split_22.xlsx")

    format_trait_matrix("genotype_evaluation/All_Traits_matrix.xlsx")
    reorder_marker_columns("formatted_matrices/All_Traits_matrix_formatted.xlsx")

    evaluate_gene_status("reordered_matrices/All_Traits_matrix_reordered.xlsx")

    insert_genotype_names("evaluated_status_matrices/All_Traits_status_evaluated.xlsx", "IMPORT/Names.xlsx")

    phenotype_function("annotated_matrices/All_Traits_matrix_annotated.xlsx")



if __name__ == "__main__":
    main()

