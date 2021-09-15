using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms; //�s�W

namespace My
{
    public class MyControl
    {
        #region �M�wTreeView�p�����`�I

        /// <summary>
        /// �M�wTreeView�p�����`�I
        /// </summary>
        /// <param name="node">��ƫ��A��TreeNode</param>
        /// <param name="FindNodeType">�n��ܸ`�I���覡
        /// Previous�@, PreviousVisible , Next , NextVisible , First , Last
        /// </param>
        public static void SelectNode(TreeNode node, string FindNodeType)
        {
            if (node.IsSelected)
            {
                // �M�wTreeNode�p��Q���.
                switch (FindNodeType)
                {
                    case "Previous":
                        node.TreeView.SelectedNode = node.PrevNode;
                        break;
                    case "PreviousVisible":
                        node.TreeView.SelectedNode = node.PrevVisibleNode;
                        break;
                    case "Next":
                        node.TreeView.SelectedNode = node.NextNode;
                        break;
                    case "NextVisible":
                        node.TreeView.SelectedNode = node.NextVisibleNode;
                        break;
                    case "First":
                        node.TreeView.SelectedNode = node.FirstNode;
                        break;
                    case "Last":
                        node.TreeView.SelectedNode = node.LastNode;
                        break;
                }
            }
            node.TreeView.Focus();
        }

        #endregion


        #region ���^���X�C�@��node����T

        /// <summary>
        /// ���^���X�C�@��node����T
        /// </summary>
        /// <param name="treeNode">�ǤJ��ƫ��A��TreeNode���ܼ�</param>
        public static void PrintRecursive(TreeNode treeNode)
        {
            // ���treeNode��T���e
            MessageBox.Show(treeNode.Text);

            // ���^���X�C�@��node����T.
            foreach (TreeNode tn in treeNode.Nodes)
            {
                PrintRecursive(tn);
            }
        }

        #endregion


        #region �NTreeView�ǤJ���{�ǨӶi��B�z.

        /// <summary>
        ///  �NTreeView�ǤJ���{�ǨӶi��B�z.
        /// </summary>
        /// <param name="treeView"></param>
        public static void CallRecursive(TreeView treeView)
        {
            // ���XTreeView���Ҧ��`�I
            TreeNodeCollection nodes = treeView.Nodes;
            foreach (TreeNode n in nodes)
            {
                PrintRecursive(n);
            }
        }

        #endregion


        #region "ComboBox�����J�ƭ�"

        public static void ComboBoxGetNumber(ComboBox comobj, int num)
        {
            int i = 0;
            comobj.Items.Clear();
            for (i = 0; i < num - 1; i++)
            {

                comobj.Items.Add(i);
            }


        }

        #endregion

    }
}
