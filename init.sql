--
-- PostgreSQL database dump
--

-- Dumped from database version 15.2 (Debian 15.2-1.pgdg110+1)
-- Dumped by pg_dump version 15.2 (Debian 15.2-1.pgdg110+1)

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: items; Type: TABLE; Schema: public; Owner: due
--

CREATE TABLE public.items (
    itemid integer NOT NULL,
    name text,
    price integer,
    ref text,
    packageid integer NOT NULL,
    warranty boolean,
    duration integer
);


ALTER TABLE public.items OWNER TO due;

--
-- Name: orders; Type: TABLE; Schema: public; Owner: due
--

CREATE TABLE public.orders (
    orderid integer NOT NULL,
    odername text NOT NULL
);


ALTER TABLE public.orders OWNER TO due;

--
-- Name: packages; Type: TABLE; Schema: public; Owner: due
--

CREATE TABLE public.packages (
    packageid integer NOT NULL,
    orderid integer NOT NULL
);


ALTER TABLE public.packages OWNER TO due;

--
-- Data for Name: items; Type: TABLE DATA; Schema: public; Owner: due
--

COPY public.items (itemid, name, price, ref, packageid, warranty, duration) FROM stdin;
\.


--
-- Data for Name: orders; Type: TABLE DATA; Schema: public; Owner: due
--

COPY public.orders (orderid, odername) FROM stdin;
\.


--
-- Data for Name: packages; Type: TABLE DATA; Schema: public; Owner: due
--

COPY public.packages (packageid, orderid) FROM stdin;
\.


--
-- Name: items items_pkey; Type: CONSTRAINT; Schema: public; Owner: due
--

ALTER TABLE ONLY public.items
    ADD CONSTRAINT items_pkey PRIMARY KEY (itemid);


--
-- Name: orders orders_pkey; Type: CONSTRAINT; Schema: public; Owner: due
--

ALTER TABLE ONLY public.orders
    ADD CONSTRAINT orders_pkey PRIMARY KEY (orderid);


--
-- Name: packages packages_pkey; Type: CONSTRAINT; Schema: public; Owner: due
--

ALTER TABLE ONLY public.packages
    ADD CONSTRAINT packages_pkey PRIMARY KEY (packageid);


--
-- Name: items items_packageid_fkey; Type: FK CONSTRAINT; Schema: public; Owner: due
--

ALTER TABLE ONLY public.items
    ADD CONSTRAINT items_packageid_fkey FOREIGN KEY (packageid) REFERENCES public.packages(packageid);


--
-- Name: packages packages_orderid_fkey; Type: FK CONSTRAINT; Schema: public; Owner: due
--

ALTER TABLE ONLY public.packages
    ADD CONSTRAINT packages_orderid_fkey FOREIGN KEY (orderid) REFERENCES public.orders(orderid);


--
-- PostgreSQL database dump complete
--
