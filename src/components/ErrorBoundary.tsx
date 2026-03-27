import React, { ErrorInfo, ReactNode } from "react";

interface Props {
  children?: ReactNode;
}

interface State {
  hasError: boolean;
}

export class ErrorBoundary extends React.Component<Props, State> {
  public state: State = {
    hasError: false
  };

  public static getDerivedStateFromError(_: Error): State {
    return { hasError: true };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
  }

  public render() {
    if (this.state.hasError) {
      return (
        <div className="p-8 bg-red-50 border border-red-200 rounded-3xl text-center">
          <h2 className="text-xl font-black text-red-600 uppercase tracking-tight mb-2">Сталася помилка</h2>
          <p className="text-red-500 font-medium">Не вдалося завантажити цей розділ. Будь ласка, спробуйте оновити сторінку.</p>
        </div>
      );
    }

    return this.props.children;
  }
}
