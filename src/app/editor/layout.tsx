export default function EditorLayout({ children }: LayoutProps<"/editor">) {
  return <div className="word-editor fixed inset-0 flex flex-col bg-word-canvas text-word-text">{children}</div>;
}
