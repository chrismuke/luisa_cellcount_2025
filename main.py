import click
import numpy as np
import openpyxl
from pathlib import Path

@click.command()
@click.option('--cells', type=click.Path(exists=True, file_okay=False, dir_okay=True), required=True, help='Directory containing the cell segmentation files.')
@click.option('--exclude', type=click.Path(exists=True, file_okay=False, dir_okay=True), required=True, help='Directory containing the exclusion region files.')
@click.option('--pixel-size', type=float, default=2.0, help='The size of a pixel in micrometers (µm).')
def analyze_cells(cells, exclude, pixel_size):
    """
    Analyzes cell segmentation files, counts cells not in excluded regions,
    and calculates cell density.
    """
    click.echo("Starting cell analysis...")
    cells_dir = Path(cells)
    exclude_dir = Path(exclude)
    pixel_area = pixel_size ** 2

    results = {}

    for cell_file in sorted(cells_dir.rglob('*_seg.npy')):
        # Ignore hidden files (like ._... on macOS)
        if cell_file.name.startswith('._'):
            continue

        relative_path = cell_file.relative_to(cells_dir)
        exclude_file = exclude_dir / relative_path

        if not exclude_file.exists():
            click.secho(f"Warning: No matching exclusion file for {cell_file}", fg='yellow')
            continue

        try:
            cell_data = np.load(cell_file, allow_pickle=True)
            exclude_data = np.load(exclude_file, allow_pickle=True)

            # The _seg.npy files from cellpose are dictionaries stored in a numpy array.
            # We extract the dictionary and then get the 'masks' key.
            if cell_data.ndim == 0:
                cell_mask = cell_data.item().get('masks')
            else:
                cell_mask = cell_data # Fallback for simple array masks

            if exclude_data.ndim == 0:
                exclude_mask = exclude_data.item().get('masks')
            else:
                exclude_mask = exclude_data # Fallback for simple array masks

            if cell_mask is None or exclude_mask is None:
                click.secho(f"Warning: 'masks' key not found or invalid format in {cell_file} or {exclude_file}", fg='yellow')
                continue

        except Exception as e:
            click.secho(f"Error processing {cell_file} or {exclude_file}: {e}", fg='red')
            continue

        # Get the top-level folder pair name for aggregation
        folder_pair_name = relative_path.parts[0]
        if folder_pair_name not in results:
            results[folder_pair_name] = {'Total Cells': 0, 'Total Area': 0}

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

        results[folder_pair_name]['Total Cells'] += valid_cell_count
        results[folder_pair_name]['Total Area'] += non_excluded_area_um2

    # After processing all files, save the aggregated results
    if results:
        save_results_to_excel(results)
    else:
        click.secho("No data processed. No output file generated.", fg='yellow')

    click.echo("\nAnalysis complete.")

def save_results_to_excel(results, output_filename="cell_count_results.xlsx"):
    """Saves the aggregated results to an Excel file."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Cell Count Analysis"

    # Write header
    header = ["Folder Pair", "Total Cells", "Total Area (µm²)", "Cell Density (cells/µm²)"]
    sheet.append(header)

    # Write data rows
    for folder_pair, data in results.items():
        total_cells = data['Total Cells']
        total_area = data['Total Area']

        if total_area > 0:
            density = total_cells / total_area
        else:
            density = 0

        row = [folder_pair, total_cells, f"{total_area:.2f}", f"{density:.4f}"]
        sheet.append(row)

    try:
        workbook.save(output_filename)
        click.secho(f"Results successfully saved to {output_filename}", fg='green')
    except Exception as e:
        click.secho(f"Error saving Excel file: {e}", fg='red')

if __name__ == '__main__':
    analyze_cells()