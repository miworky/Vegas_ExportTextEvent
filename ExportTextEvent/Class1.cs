using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

using System.Drawing;

namespace ExportTextEvent
{ 
    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            if (vegas.Project.Tracks.Count == 0)
            {
                // トラックがないときは何もしない
                return;
            }

            // ダイアログを開き出力する Text のファイルパスをユーザーに選択させる
            string saveFilePath = GetFilePath(vegas.Project.FilePath, "ExportTextEvent");
            if (saveFilePath.Length == 0)
            {
                return;
            }

            // タイムライン上にある Video トラックの静止画・動画のファイル名を全部集める
            List<Tuple<long, string>> textEvents = new List<Tuple<long, string>>();
            foreach (Track track in vegas.Project.Tracks)
            {
                foreach (TrackEvent trackEvent in track.Events)
                {
                    if (!trackEvent.IsVideo())
                    {
                        // ビデオトラック以外は無視する
                        continue;
                    }

                    // アクティブテイクのみを対象にする
                    Take take = trackEvent.ActiveTake;
                    if (take == null)
                    {
                        // アクティブテイクがなければ無視
                        continue;
                    }

                    // テキストイベントのテキストを取得する
                    string planeText = GetPlaneTextFromMedia(take.Media);
                    if (planeText.Length == 0)
                    {
                        continue;
                    }

                    // テキストが見つかった
                    // このファイルが張り付けられているフレーム位置とテキストのペアを追加する
                    textEvents.Add(Tuple.Create(trackEvent.Start.FrameCount, planeText));
                }
            }

            // TimeCodeで昇順にソート
            textEvents.Sort( (a, b) => (int)(a.Item1 - b.Item1)  );

            // 出力用のファイルを開く
            System.IO.StreamWriter writer = new System.IO.StreamWriter(saveFilePath, false, Encoding.GetEncoding("Shift_JIS"));

            // ファイルに出力する
            Timecode timecode = new Timecode();
            foreach (var textEvent in textEvents)
            {
                timecode.FrameCount = textEvent.Item1;
                string text = textEvent.Item2;
                string textWithoutNewLine = text.Replace('\n', ' ');    // 改行はスペースに置き換える

                writer.WriteLine(timecode.ToString() + " " + textWithoutNewLine);
            }

            writer.Close();

            MessageBox.Show("終了しました。");
        }


        // ダイアログを開きファイルパスをユーザーに選択させる
        private string GetFilePath(string rootFilePath, string preFix)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = preFix + System.IO.Path.GetFileNameWithoutExtension(rootFilePath) + ".txt";
            sfd.InitialDirectory = System.IO.Path.GetDirectoryName(rootFilePath) + "\\";
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return "";
            }

            return sfd.FileName;
        }

        private OFXEffect GetOFXEffect(Media media)
        {
            if (media == null)
            {
                return null;
            }

            Effect generator = media.Generator;
            if (generator == null)
            {
                return null;
            }

            OFXEffect ofxEffect = generator.OFXEffect;
            if (ofxEffect == null)
            {
                return null;
            }

            return ofxEffect;
        }

        private string GetPlaneTextFromMedia(Media media)
        {
            OFXEffect ofxEffect = GetOFXEffect(media);
            if (ofxEffect == null)
            {
                return "";
            }

            string planeText = GetPlaneTextFromTextEvent(ofxEffect);
            if (planeText.Length == 0)
            {
                return "";
            }

            return planeText;
        }

        private string GetPlaneTextFromTextEvent(OFXEffect ofxEffect)
        {
            OFXStringParameter textParam = ofxEffect.FindParameterByName("Text") as OFXStringParameter;
            if (textParam == null)
            {
                // TextEvent でなければ無視
                return "";
            }

            string rtfData = textParam.Value;   // rtf形式のテキスト

            RichTextBox richtextBox = new RichTextBox();
            richtextBox.Rtf = rtfData;

            string planeText = richtextBox.Text;
            return planeText;
        }


    }
}