import marimo

__generated_with = "0.11.13"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    return (mo,)


@app.cell
def _():
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import tqdm
    import py

    return np, pd, plt, py, tqdm


@app.cell
def _(pd, py):

    dataset_root = py.path.local("calls_snapshot_dataset")
    paths = []
    for img in dataset_root.visit("*.jpg"):
        paths.append(img.relto(dataset_root))

    df = pd.DataFrame(paths, columns=["screenshot_path"])
    return dataset_root, df, img, paths


@app.cell
def _(df):
    df['account'] = df.screenshot_path.apply(lambda x: x.split("/")[0].replace("account=",""))
    df['gong_call_id'] = df.screenshot_path.apply(lambda x: x.split("/")[1])
    df['gong_call_link'] = "https://gong.app.gong.io/call?id=" +  df['gong_call_id']
    return


@app.cell
def _(dataset_root):
    from paddleocr import PaddleOCR

    ocr = PaddleOCR(use_angle_cls=True, lang='en')

    def extract_ocr(path):
        result = ocr.ocr(str(dataset_root / path), cls=True)
        text = ""
        if result and result[0]:
            for line in result[0]:
                text += (line[1][0]) + "\n"
        return text
    return PaddleOCR, extract_ocr, ocr


@app.cell
def _(df, extract_ocr, tqdm):
    tqdm.tqdm().pandas()

    df['ocr_extracted_text'] = df.screenshot_path.progress_apply(extract_ocr)
    return


@app.cell
def _(df):
    df['ocr_extracted_text_len'] = df.ocr_extracted_text.str.len()
    return


@app.cell
def _():
    import seaborn as sns
    return (sns,)


@app.cell
def _(df, sns):
    sns.scatterplot(df, x='account', y='ocr_extracted_text_len')
    return


@app.cell
def _():
    from transformers import BertTokenizer, BertModel
    import torch

    # Load pre-trained BERT tokenizer and model
    tokenizer = BertTokenizer.from_pretrained('bert-base-uncased')
    model = BertModel.from_pretrained('bert-base-uncased')

    def bert_embeddings(text):
        # Tokenize input
        inputs = tokenizer(text, return_tensors="pt")
    
        # Pass through model
        with torch.no_grad():
            outputs = model(**inputs)
    
        token_embeddings = outputs.last_hidden_state  # [batch_size, seq_len, hidden_size]
    
        # Take the mean over the token embeddings
        sentence_embedding = token_embeddings.mean(dim=1)  # [batch_size, hidden_size]
    
        return sentence_embedding.squeeze(0).numpy()  # remove batch dimension

    return BertModel, BertTokenizer, bert_embeddings, model, tokenizer, torch


@app.cell
def _(df):
    from sentence_transformers import SentenceTransformer

    # Load a model
    model_ = SentenceTransformer('all-MiniLM-L6-v2')

    def sbert_embeddings(text):
        embedding = model_.encode(text, convert_to_numpy=True)
        return embedding

    df['bert_embeddings'] = df.ocr_extracted_text.progress_apply(sbert_embeddings)

    return SentenceTransformer, model_, sbert_embeddings


@app.cell
def _(bert_embeddings, df):
    bert_embeddings(df.iloc[0].ocr_extracted_text)
    return


@app.cell
def _(bert_embeddings, df):
    df['bert_embeddings'] = df.ocr_extracted_text.progress_apply(bert_embeddings)
    return


@app.cell
def _(df, np, plt, sns):
    from sklearn.preprocessing import normalize
    from sklearn.metrics.pairwise import cosine_similarity

    embeddings = np.vstack(df['bert_embeddings'].values)  # shape (num_slides, embedding_size)
    embeddings = normalize(embeddings)  # L2 normalization

    # Compute cosine similarity matrix
    similarity_matrix = cosine_similarity(embeddings)

    plt.figure(figsize=(12, 12))
    plt.title("Similiary Heatmap (Bert embeddings)")
    sns.heatmap(similarity_matrix)
    return cosine_similarity, embeddings, normalize, similarity_matrix


@app.cell
def _(np):
    def find_top_k_similar(current_index, similarity_matrix, df, k=3):
        sim_scores = similarity_matrix[current_index]
        # Exclude self (similarity 1.0 at own index)
        top_indices = np.argsort(sim_scores)[::-1] 
        results = []
        idx = 0
        while k >= 0:
            if top_indices[idx] != current_index:
                sim_row = df.loc[top_indices[idx]]
                results.append({
                    'index': top_indices[idx],
                    'other_account': sim_row['account'],
                    'similar_screenshot_path': sim_row['screenshot_path'],
                    'similar_score': sim_scores[idx]
                })
            idx+=1
            k -= 1
        return results

    return (find_top_k_similar,)


@app.cell
def _(df, find_top_k_similar, similarity_matrix):
    df['top_3_similar_slides'] = df.index.to_series().apply(lambda idx: find_top_k_similar(idx, similarity_matrix, df))
    return


@app.cell
def _():
    return


@app.cell
def _(dataset_root, df, plt):
    from PIL import Image

    def show_similar_slides(index):
        """
        Display the main slide and its top-3 similar slides.
    
        Args:
            row: A row from the DataFrame.
            base_path: Optional base directory if paths are relative.
        """
        row = df.iloc[index]
        main_path = dataset_root / row['screenshot_path']
        top_similar = row['top_3_similar_slides']

        fig, axes = plt.subplots(1, 4, figsize=(20, 5))
        fig.suptitle('Main Slide and Top-3 Similar Slides', fontsize=16)

        # Plot main image
        try:
            img =  Image.open(main_path)
            axes[0].imshow(img)
            axes[0].set_title(f"Main Slide| #{index}")
        except Exception as e:
            print(f"Error loading main image: {e}")
            axes[0].text(0.5, 0.5, 'Failed to load', ha='center', va='center')
        axes[0].axis('off')

        # Plot similar images
        for i, similar_info in enumerate(top_similar):
            try:
                sim_path = dataset_root / similar_info['similar_screenshot_path']
                sim_score = similar_info['similar_score']
                img = Image.open(sim_path)
                axes[i + 1].imshow(img)
                axes[i + 1].set_title(f"Score: {sim_score:.2f} | #{similar_info['index']}")
            except Exception as e:
                print(f"Error loading similar image {i}: {e}")
                axes[i + 1].text(0.5, 0.5, 'Failed to load', ha='center', va='center')
            axes[i + 1].axis('off')

        plt.tight_layout()
        plt.show()

    return Image, show_similar_slides


@app.cell
def _(show_similar_slides):
    show_similar_slides(46)
    return


@app.cell
def _(df):
    # Helper function to get distinct accounts and their highest similarity scores
    def get_distinct_accounts_and_scores(top_k_similar_slides):
        account_scores = {}
    
        # Iterate through each similar slide and track the highest score for each account
        for slide in top_k_similar_slides:
            account = slide['other_account']
            score = slide['similar_score']
        
            # If the account is not in the dictionary, add it
            if account not in account_scores:
                account_scores[account] = score
            else:
                # If it is already in the dictionary, take the higher score
                account_scores[account] = min(account_scores[account], score)
    
        # Return the number of distinct accounts and the sum of their highest similarity scores
        return len(account_scores), sum(account_scores.values())

    # Apply the function to each row's 'top_3_similar_slides'
    df['distinct_accounts_count'], df['top_similarity_score_sum'] = zip(*df['top_3_similar_slides'].apply(get_distinct_accounts_and_scores))

    # Sort by distinct accounts count (descending), and if tied, by the sum of similarity scores (descending)
    df_sorted_by_diversity_and_score = df[df['distinct_accounts_count']> 1].sort_values(by=['top_similarity_score_sum', 'distinct_accounts_count'], ascending=[True, False])

    # Select the top 4 rows with the most distinct accounts and high similarity scores
    top_4_diverse_and_high_score_rows = df_sorted_by_diversity_and_score.head(20)

    # Display the top 4 rows with their distinct account count and similarity score sum
    top_4_diverse_and_high_score_rows[['account', 'top_3_similar_slides', 'distinct_accounts_count', 'top_similarity_score_sum']]

    return (
        df_sorted_by_diversity_and_score,
        get_distinct_accounts_and_scores,
        top_4_diverse_and_high_score_rows,
    )


@app.cell
def _(show_similar_slides):
    show_similar_slides(0)
    return


@app.cell
def _(df):
    def print_side_by_side(*texts, width=40, padding=4):
        # First split each text into lines
        split_texts = [text.splitlines() for text in texts]
        # Find the maximum number of lines among all texts
        max_lines = max(len(lines) for lines in split_texts)
        # Pad each text with empty lines if needed
        split_texts = [lines + [''] * (max_lines - len(lines)) for lines in split_texts]
    
        for line_set in zip(*split_texts):
            # For each line across texts, format them to fixed width and join
            print((' ' * padding).join(line.ljust(width) for line in line_set))


    # Example usage:
    print_side_by_side(
        df.iloc[0].ocr_extracted_text,
        df.iloc[5].ocr_extracted_text,
        df.iloc[9].ocr_extracted_text,
        width=50,  # Adjust based on your text
        padding=6  # Adjust spacing between columns
    )

    return (print_side_by_side,)


if __name__ == "__main__":
    app.run()
