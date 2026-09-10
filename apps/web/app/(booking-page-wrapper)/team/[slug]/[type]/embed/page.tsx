import { loadTranslations } from "@calcom/i18n/server";
import TypePage, { type PageProps as ClientPageProps } from "@calcom/web/modules/team/type-view";
import { buildLegacyCtx } from "@lib/buildLegacyCtx";
import { getServerSideProps } from "@lib/team/[slug]/[type]/getServerSideProps";
import type { PageProps } from "app/_types";
import { CustomI18nProvider } from "app/CustomI18nProvider";
import withEmbedSsrAppDir from "app/WithEmbedSSR";
import { cookies, headers } from "next/headers";

const getData = withEmbedSsrAppDir<ClientPageProps>(getServerSideProps);

export const generateMetadata = async () => ({ robots: { follow: false, index: false } });

export default async function TeamBookingEmbedPage({ params, searchParams }: PageProps) {
  const context = buildLegacyCtx(await headers(), await cookies(), await params, await searchParams);
  const props = await getData(context);
  const locale = props.eventData?.interfaceLanguage;

  if (locale) {
    const ns = "common";
    const translations = await loadTranslations(locale, ns);
    return (
      <CustomI18nProvider translations={translations} locale={locale} ns={ns}>
        <TypePage {...props} />
      </CustomI18nProvider>
    );
  }

  return <TypePage {...props} />;
}
