
import type { SavedPhoto } from '@/lib/photos';
import {
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  TextRun,
} from 'docx';
import * as FileSystem from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';
import { Image } from 'react-native';

const FILE_NAME = 'QuickCamNotes-Export.docx';

export type ExportProgress = {
  current: number;
  total: number;
  stage: 'photos' | 'document' | 'saving' | 'sharing';
};

function getImageDimensions(
  uri: string
): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    Image.getSize(
      uri,
      (width, height) => resolve({ width, height }),
      reject
    );
  });
}

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes;
}

export async function exportPhotosToWord(
  photos: SavedPhoto[],
  onProgress?: (progress: ExportProgress) => void
): Promise<void> {
  if (photos.length === 0) {
    throw new Error('No photos available to export.');
  }

  const total = photos.length;

  const report =  async (
  current: number,
  stage: ExportProgress['stage']
) => {
  onProgress?.({ current, total, stage });

  // Allow the UI to render the updated progress.
  await new Promise(resolve => setTimeout(resolve, 100));
};

  const sections: Paragraph[] = [
    new Paragraph({
      text: 'QuickCamNotes',
      heading: HeadingLevel.TITLE,
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Exported ${total} photo(s)`,
          color: '666666',
        }),
      ],
      spacing: { after: 400 },
    }),
  ];

  await report(0, 'photos');

  for (let index = 0; index < total; index++) {
    const photo = photos[index];

    const base64 = await FileSystem.readAsStringAsync(photo.uri, {
      encoding: FileSystem.EncodingType.Base64,
    });

    const imageData = decodeBase64(base64);
    const dimensions = await getImageDimensions(photo.uri);

    const maxWidth = 500;
    const maxHeight = 600;

    const scale = Math.min(
      maxWidth / dimensions.width,
      maxHeight / dimensions.height,
      1
    );

    const width = Math.round(dimensions.width * scale);
    const height = Math.round(dimensions.height * scale);

    sections.push(
      new Paragraph({
        text: `Photo ${index + 1}`,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 300, after: 150 },
      }),
      new Paragraph({
        children: [
          new ImageRun({
            data: imageData,
            transformation: { width, height },
            type: 'jpg',
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Note: ', bold: true }),
          new TextRun(photo.note || 'No note added.'),
        ],
        spacing: { after: 120 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Saved on: ', bold: true }),
          new TextRun(new Date(photo.createdAt).toLocaleString()),
        ],
        spacing: { after: 300 },
      })
    );

    await report(index + 1, 'photos');
  }

  await report(total, 'document');

  const document = new Document({
    sections: [{ children: sections }],
  });

  const base64Document = await Packer.toBase64String(document);

  await report(total, 'saving');

  const fileUri = `${FileSystem.cacheDirectory}${FILE_NAME}`;

  await FileSystem.writeAsStringAsync(fileUri, base64Document, {
    encoding: FileSystem.EncodingType.Base64,
  });

  await report(total, 'sharing');

  const available = await Sharing.isAvailableAsync();

  if (!available) {
    throw new Error('Sharing is not available on this device.');
  }

  await Sharing.shareAsync(fileUri, {
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    dialogTitle: 'Export QuickCamNotes',
    UTI: 'org.openxmlformats.wordprocessingml.document',
  });
}