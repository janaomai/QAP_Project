import matplotlib.pyplot as plt

def plot_z_scores(site, scores, img_stream):
    all_cycles = ["EQA2401", "EQA2402", "EQA2403", "EQA2404", "EQA2405", "EQA2406", "EQA2407", "EQA2408", "EQA2409", "EQA2410", "EQA2411", "EQA2412"]

    fig, ax = plt.subplots(figsize=(5, 4))

    # Extract cycles and corresponding Z-scores, filtering out "No submission"
    cycles = []
    z_values = []
    for cycle, z_score in scores.items():
        if z_score != "No submission":
            cycles.append(all_cycles.index(cycle))
            z_values.append(z_score)

    ax.plot(cycles, z_values, 'ko-', markersize=8)  # Plot the Z-scores with lines and markers

    # Set y-axis limits and labels
    ax.set_ylim(-3, 3)
    ax.set_yticks([-3, -2, -1, 0, 1, 2, 3])
    ax.set_yticklabels([f"{x:+d}" if x != 0 else "0" for x in [-3, -2, -1, 0, 1, 2, 3]])  # Include the '+' sign for positive numbers and '-' for negative numbers

    # Add transparent rectangles
    ax.axhspan(-3, -2, facecolor='red', alpha=0.5)
    ax.axhspan(-2, -1, facecolor='red', alpha=0.3)
    ax.axhspan(1, 2, facecolor='red', alpha=0.3)
    ax.axhspan(2, 3, facecolor='red', alpha=0.5)

    # Set x-axis labels
    ax.set_xticks(range(len(all_cycles)))
    ax.set_xticklabels(all_cycles, rotation=45)

    ax.set_title(f"{site}'s QAP Performance", fontsize=10)
    ax.set_xlabel("Cycle", fontsize=10)
    ax.set_ylabel("SD", fontsize=10)

    plt.tight_layout()
    plt.savefig(img_stream, format='png')
    plt.close()
