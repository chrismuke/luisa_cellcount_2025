import click
import numpy as np
import openpyxl
from pathlib import Path
from datetime import datetime


def parse_path_to_columns(relative_path, column_names):
    """
    Parse a relative path into column values based on column names.

    The path parts (folders + filename without _seg.npy) are mapped to column names.
    Example: HS2ST1/14684/Zentral/Snap-10908-Image Export-07_c1-3_seg.npy
    with columns=['gene', 'mouse', 'region', 'image'] produces:
    {'gene': 'HS2ST1', 'mouse': '14684', 'region': 'Zentral', 'image': 'Snap-10908-Image Export-07_c1-3'}
    """
    parts = list(relative_path.parent.parts)
    # Add filename without _seg.npy suffix
    filename = relative_path.name
    if filename.endswith('_seg.npy'):
        filename = filename[:-8]
    parts.append(filename)

    result = {}
    for i, col_name in enumerate(column_names):
        if i < len(parts):
            result[col_name] = parts[i]
        else:
            result[col_name] = ''

    return result


@click.command()
@click.option('--cells', type=click.Path(exists=True, file_okay=False, dir_okay=True), required=True, help='Directory containing the cell segmentation files.')
@click.option('--exclude', type=click.Path(exists=True, file_okay=False, dir_okay=True), required=True, help='Directory containing the exclusion region files.')
@click.option('--pixel-size', type=float, default=2.0, help='The size of a pixel in micrometers (µm).')
@click.option('--columns', type=str, default=None, help='Comma-separated column names derived from path (e.g., "gene,mouse,region,image"). Maps to folder1/folder2/.../filename.')
def analyze_cells(cells, exclude, pixel_size, columns):
    """
    Analyzes cell segmentation files, counts cells not in excluded regions,
    and calculates cell density.
    """
    click.echo("Starting cell analysis...")
    cells_dir = Path(cells)
    exclude_dir = Path(exclude)
    pixel_area = pixel_size ** 2

    # Parse column names if provided
    column_names = [c.strip() for c in columns.split(',')] if columns else None

    results = []

    for cell_file in sorted(cells_dir.rglob('*_seg.npy')):
        # Ignore hidden files (like ._... on macOS)
        if cell_file.name.startswith('._'):
            continue

        relative_path = cell_file.relative_to(cells_dir)
        exclude_file = exclude_dir / relative_path

        try:
            cell_data = np.load(cell_file, allow_pickle=True)

            # The _seg.npy files from cellpose are dictionaries stored in a numpy array.
            # We extract the dictionary and then get the 'masks' key.
            if cell_data.ndim == 0:
                cell_mask = cell_data.item().get('masks')
            else:
                cell_mask = cell_data  # Fallback for simple array masks

            if cell_mask is None:
                click.secho(f"Warning: 'masks' key not found in {cell_file}", fg='yellow')
                continue

            # Load exclusion mask if it exists, otherwise use zeros (no exclusions)
            if exclude_file.exists():
                exclude_data = np.load(exclude_file, allow_pickle=True)
                if exclude_data.ndim == 0:
                    exclude_mask = exclude_data.item().get('masks')
                else:
                    exclude_mask = exclude_data  # Fallback for simple array masks

                if exclude_mask is None:
                    click.secho(f"Warning: 'masks' key not found in {exclude_file}, treating as no exclusion", fg='yellow')
                    exclude_mask = np.zeros_like(cell_mask)
            else:
                # No exclusion file means no regions to exclude
                exclude_mask = np.zeros_like(cell_mask)

        except Exception as e:
            click.secho(f"Error processing {cell_file}: {e}", fg='red')
            continue

        # Find unique cell IDs (excluding 0, which is the background)
        unique_cells = np.unique(cell_mask[cell_mask > 0])
        valid_cell_count = 0

        for cell_id in unique_cells:
            # Find the coordinates of the current cell
            cell_coords = np.argwhere(cell_mask == cell_id)
            # Calculate the center of the cell
            center = cell_coords.mean(axis=0)
            center_y, center_x = int(center[0]), int(center[1])

            # Check if the center is in an excluded region (mask value > 0)
            if exclude_mask[center_y, center_x] == 0:
                valid_cell_count += 1

        # Calculate the non-excluded area
        non_excluded_area_pixels = np.sum(exclude_mask == 0)
        non_excluded_area_um2 = non_excluded_area_pixels * pixel_area

        # Calculate cell density for this file
        if non_excluded_area_um2 > 0:
            density = valid_cell_count / non_excluded_area_um2
        else:
            density = 0

        # Build result dict
        result = {
            'total_cells': valid_cell_count,
            'total_area': non_excluded_area_um2,
            'cell_density': density
        }

        # Add path-based columns if specified
        if column_names:
            path_columns = parse_path_to_columns(relative_path, column_names)
            result['path_columns'] = path_columns
        else:
            # Fallback to original behavior
            parent_folder = relative_path.parent.name if relative_path.parent != Path('.') else ''
            result['file_path'] = str(relative_path)
            result['parent_folder'] = parent_folder

        results.append(result)

    # After processing all files, save the results
    if results:
        # Construct output filename from parent folder + cells folder name + pixel size + timestamp
        parent_name = cells_dir.parent.name.replace(' ', '_')
        cells_name = cells_dir.name.replace(' ', '_')
        pixel_size_str = f"{pixel_size:.2f}".replace('.', 'p')
        timestamp = datetime.now().strftime('%Y-%m-%d_%H_%M_%S')
        output_filename = f"{parent_name}_{cells_name}_px{pixel_size_str}_{timestamp}.xlsx"
        save_results_to_excel(results, output_filename, column_names)
    else:
        click.secho("No data processed. No output file generated.", fg='yellow')

    click.echo("\nAnalysis complete.")


def save_results_to_excel(results, output_filename="cell_count_results.xlsx", column_names=None):
    """Saves the per-file results to an Excel file."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Cell Count Analysis"

    # Write header
    if column_names:
        header = column_names + ["Total Cells", "Total Area (µm²)", "Cell Density (cells/µm²)"]
    else:
        header = ["Parent Folder", "File Path", "Total Cells", "Total Area (µm²)", "Cell Density (cells/µm²)"]
    sheet.append(header)

    # Write data rows for each file
    for file_result in results:
        total_cells = file_result['total_cells']
        total_area = file_result['total_area']
        density = file_result['cell_density']

        if column_names:
            path_cols = file_result['path_columns']
            row = [path_cols.get(col, '') for col in column_names]
        else:
            row = [file_result['parent_folder'], file_result['file_path']]

        row.extend([total_cells, f"{total_area:.2f}", f"{density:.6f}"])
        sheet.append(row)

    try:
        workbook.save(output_filename)
        click.secho(f"Results successfully saved to {output_filename}", fg='green')
    except Exception as e:
        click.secho(f"Error saving Excel file: {e}", fg='red')


if __name__ == '__main__':
    analyze_cells()
